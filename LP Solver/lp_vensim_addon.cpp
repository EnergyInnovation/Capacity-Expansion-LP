// lp_vensim_addon.cpp — HiGHS C API (multi-scenario + storage + K-block dispatch + MPPC fast quantize)
// Version code must remain 62051.
//
// Features retained & integrated:
//  • Region-first layout; Build, Retire, Dispatch (+K merit blocks per tech using NSDoDC[E])
//  • Reserve margin enforced on peak slices; imports count; storage contributes via capacity credit
//  • Trading with directional caps TransRR[R,R] (<=0 disables direction)
//  • Region-specific costs: CAPEX[R,E], FOM_new[R,E], Fuel[R,E], VOM[R,E]
//  • CES / RPS with optional ACP slacks
//  • Per-(R,E) MaxBuild / CapMax
//  • Single storage tech per region (4-hour only): build power (MW), SoC (MWh), charge/discharge (MWh)
//  • SoC dynamics; cyclic or SoC0-based initialization; power/energy limits; charge/discharge VOM
//  • Multi-scenario snapshots and getters (no renames)
//  • LMPs (demand duals), reserve duals (peak-only, clamped), revenues, flows
//  • Revenue per MW installed (post-decision), capacity price ($/MW-yr)
//  • MPPC fast approach (post-quantize in getters) using MPPC_MinSize[E] and switch
//
// Build (x64 Native Tools for VS):
//   del /q lp_vensim_addon.obj lp_vensim_addon.lib lp_vensim_addon.exp lp_vensim_addon.dll
//   cl /nologo /LD /O2 /EHsc /MT ^
//     "lp_vensim_addon.cpp" ^
//     /I "C:\\vcpkg\\installed\\x64-windows-static\\include" ^
//     /I "C:\\vcpkg\\installed\\x64-windows-static\\include\\highs" ^
//     /link /NOLOGO ^
//     /LIBPATH:"C:\\vcpkg\\installed\\x64-windows-static\\lib" highs.lib zlib.lib ^
//     /OUT:"lp_vensim_addon.dll"
// If highs.lib is /MD, switch /MT -> /MD and use non-*-static triplet paths.

#define NOMINMAX
#include <windows.h>
#include <vector>
#include <mutex>
#include <algorithm>
#include <cstdint>
#include <cmath>
#include <unordered_map>
#include <limits>
#include <highs/interfaces/highs_c_api.h>

#if defined(_MSC_VER)
  #define VEFCC __stdcall
#else
  #define VEFCC
#endif

using std::vector;

// ===== Vensim ABI =====
typedef double COMPREAL;
typedef struct { COMPREAL* vals; const COMPREAL* firstval; const void* dim_info; const char* varname; } VECTOR_ARG;
typedef union  { COMPREAL val; VECTOR_ARG* vec; void* tab; char* literal; void* constmat; void* datamat; } VV;

// ===== Globals (current solve) =====
static std::mutex     g_mutex;
static vector<double> g_solution, g_cost_cap, g_cost_energy;
static int            g_status=-1, g_lastCode=-999;
static double         g_obj=0.0;
static bool           g_solved=false;

// Shape / layout
static int g_R=0,g_S=0,g_H=0,g_E=0,g_T=0;
static int g_colsPerReg=0,g_N_region=0,g_N_total=0;
static int g_colSlackCES=-1,g_colSlackRPS=-1;
static bool g_haveTrade=false;

// K-block dispatch
static int  g_K=1;           // blocks per tech
static bool g_anchorMean=1;  // anchor to mean toggle

// Reporting (last solve)
static vector<double> g_lmp;            // [R*T]
static vector<double> g_capdual;        // [R*T], peak-only, clamped >=0
static vector<double> g_rev_tech;       // [R*E] energy revenue $/yr
static vector<double> g_hours_SH;       // [S*H]
static vector<double> g_cf_RESH;        // [R,E,S,H]
static vector<double> g_exist_cap;      // [R,E] existing gen (MW)

// Storage reporting comes from solution by indices

// MPPC & NSDoDC caches for getters
static vector<double> g_mppc_E;         // [E] minimum plant size (MW)
static int            g_mppc_switch=0;  // 0 off, 1 on
static vector<double> g_nsd_E;          // [E] normalized std dev of dispatch costs

// ===== Multi-scenario snapshot =====
struct Context{
  int R=0,S=0,H=0,E=0,T=0,K=1;
  int colsPerReg=0,N_region=0,N_total=0;
  int colSlackCES=-1,colSlackRPS=-1;
  bool haveTrade=false;
  bool anchorMean=true;

  vector<double> solution, cost_cap, cost_energy;
  int status=-1,lastCode=-999;
  double obj=1e308;
  bool solved=false;

  vector<double> lmp, capdual, rev_tech;
  vector<double> hours_SH, cf_RESH, exist_cap;
  vector<double> mppc_E, nsd_E;
  int mppc_switch=0;

  // Storage inputs we need for reserve RHS adjustments
  vector<double> storExistPower_R; // [R] MW
};
static std::unordered_map<int,Context> g_ctx;

static inline void save_context(int sid){
  Context c;
  c.R=g_R; c.S=g_S; c.H=g_H; c.E=g_E; c.T=g_T; c.K=g_K;
  c.colsPerReg=g_colsPerReg; c.N_region=g_N_region; c.N_total=g_N_total;
  c.colSlackCES=g_colSlackCES; c.colSlackRPS=g_colSlackRPS; c.haveTrade=g_haveTrade; c.anchorMean=g_anchorMean;
  c.solution=g_solution; c.cost_cap=g_cost_cap; c.cost_energy=g_cost_energy;
  c.status=g_status; c.lastCode=g_lastCode; c.obj=g_obj; c.solved=g_solved;
  c.lmp=g_lmp; c.capdual=g_capdual; c.rev_tech=g_rev_tech;
  c.hours_SH=g_hours_SH; c.cf_RESH=g_cf_RESH; c.exist_cap=g_exist_cap;
  c.mppc_E=g_mppc_E; c.nsd_E=g_nsd_E; c.mppc_switch=g_mppc_switch;
  g_ctx[sid]=std::move(c);
}
static inline const Context* get_ctx(int sid){ auto it=g_ctx.find(sid); return it==g_ctx.end()? nullptr : &it->second; }

// ===== Indices =====
static inline int idx_dem  (int r,int s,int h,int S,int H){ return ((r*S)+s)*H + h; }              // [R,S,H]
static inline int idx_hours(int s,int h,int H){ return s*H + h; }                                    // [S,H]
static inline int idx_cf   (int r,int e,int s,int h,int E,int S,int H){ return (((r*E+e)*S+s)*H+h);} // [R,E,S,H]
static inline int idx_mask (int s,int h,int H){ return s*H + h; }                                    // [S,H]
static inline int idx_re   (int r,int e,int E){ return r*E + e; }                                    // [R,E]

// === Region block layout (with K blocks and storage) ===
//
// Per region columns in order:
//   [Build E][Retire E][Gen (E*K*T)][Flows (R-1)*T if any][StorBuild 1][StorCharge T][StorDischarge T][StorSoC T]
//
static inline int offset_build(int /*E*/){ return 0; }
static inline int col_build_e(int e){ return e; }                     // 0..E-1
static inline int col_retire_e(int E,int e){ return E + e; }          // E..2E-1

static inline int offset_gen(int E){ return 2*E; }
static inline int col_gen_e_k_tau(int E,int K,int e,int k,int tau,int T){
  return offset_gen(E) + ((e*K + k)*T + tau);
}

static inline int offset_trade(int E,int K,int T){ return 2*E + (E*K*T); }
static inline int dest_idx_for_origin(int r,int d){ return (d<r)? d : (d-1); }
static inline int col_flow_od_tau(int E,int K,int T,int R,int r,int d,int tau){
  return offset_trade(E,K,T) + dest_idx_for_origin(r,d)*T + tau;
}

static inline int offset_stor_build(int E,int K,int T,int R,bool haveTrade){
  return offset_trade(E,K,T) + (haveTrade ? (R-1)*T : 0);
}
static inline int col_stor_build(int E,int K,int T,int R,bool haveTrade){
  return offset_stor_build(E,K,T,R,haveTrade);
}
static inline int offset_stor_charge(int E,int K,int T,int R,bool haveTrade){
  return offset_stor_build(E,K,T,R,haveTrade) + 1;
}
static inline int col_stor_charge_tau(int E,int K,int T,int R,bool haveTrade,int tau){
  return offset_stor_charge(E,K,T,R,haveTrade) + tau;
}
static inline int offset_stor_discharge(int E,int K,int T,int R,bool haveTrade){
  return offset_stor_charge(E,K,T,R,haveTrade) + T;
}
static inline int col_stor_discharge_tau(int E,int K,int T,int R,bool haveTrade,int tau){
  return offset_stor_discharge(E,K,T,R,haveTrade) + tau;
}
static inline int offset_stor_soc(int E,int K,int T,int R,bool haveTrade){
  return offset_stor_discharge(E,K,T,R,haveTrade) + T;
}
static inline int col_stor_soc_tau(int E,int K,int T,int R,bool haveTrade,int tau){
  return offset_stor_soc(E,K,T,R,haveTrade) + tau;
}

static inline int cols_per_region(int R,int E,int K,int T,bool haveTrade){
  return 2*E + (E*K*T) + (haveTrade ? (R-1)*T : 0) + 1 + 3*T;
}

static inline void reset_cache(int R,int E,int T,bool haveCES,bool haveRPS,bool haveTrade,int K,bool anchor){
  g_haveTrade   = haveTrade;
  g_K           = (K<1?1:K);
  g_anchorMean  = anchor?true:false;

  g_colsPerReg  = cols_per_region(R,E,g_K,T,haveTrade);
  g_N_region    = R*g_colsPerReg;
  g_N_total     = g_N_region + (haveCES?1:0) + (haveRPS?1:0);

  g_solution.assign(g_N_total,0.0);
  g_cost_cap.assign(R,0.0);
  g_cost_energy.assign(R,0.0);
  g_status=-1; g_lastCode=-999; g_obj=0.0; g_solved=false;
  g_R=R; g_E=E; g_T=T;

  int off = g_N_region;
  g_colSlackCES = haveCES? off++ : -1;
  g_colSlackRPS = haveRPS? off++ : -1;

  g_lmp.clear(); g_capdual.clear(); g_rev_tech.clear();
}

// ===== Triplet -> CSC =====
struct Triplet{ int row; int col; double val; };
static void triplets_to_csc(int nRows,int nCols,const vector<Triplet>& t, vector<int>& astart, vector<int>& aindex, vector<double>& avalue){
  astart.assign(nCols+1,0);
  for(const auto& x:t) if(x.col>=0 && x.col<nCols) astart[x.col+1]++;
  for(int c=0;c<nCols;++c) astart[c+1]+=astart[c];
  int nnz=(int)t.size();
  aindex.assign(nnz,0); avalue.assign(nnz,0.0);
  vector<int> next=astart;
  for(const auto& x:t){ int p=next[x.col]++; aindex[p]=x.row; avalue[p]=x.val; }
}

// ===== Normal quantile (Acklam approx) =====
static inline double inv_norm_cdf(double p){
  // clamp
  if (p<=0.0) return -1e3;
  if (p>=1.0) return  1e3;
  // Coeffs
  static const double a1=-3.969683028665376e+01, a2= 2.209460984245205e+02,
                      a3=-2.759285104469687e+02, a4= 1.383577518672690e+02,
                      a5=-3.066479806614716e+01, a6= 2.506628277459239e+00;
  static const double b1=-5.447609879822406e+01, b2= 1.615858368580409e+02,
                      b3=-1.556989798598866e+02, b4= 6.680131188771972e+01,
                      b5=-1.328068155288572e+01;
  static const double c1=-7.784894002430293e-03, c2=-3.223964580411365e-01,
                      c3=-2.400758277161838e+00, c4=-2.549732539343734e+00,
                      c5= 4.374664141464968e+00, c6= 2.938163982698783e+00;
  static const double d1= 7.784695709041462e-03, d2= 3.224671290700398e-01,
                      d3= 2.445134137142996e+00, d4= 3.754408661907416e+00;
  const double plow  = 0.02425;
  const double phigh = 1.0 - plow;
  double q,r;
  if (p < plow) {
    q = std::sqrt(-2*std::log(p));
    return (((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6) /
           ((((d1*q+d2)*q+d3)*q+d4)*q+1);
  }
  if (phigh < p) {
    q = std::sqrt(-2*std::log(1-p));
    return -(((((c1*q+c2)*q+c3)*q+c4)*q+c5)*q+c6) /
             ((((d1*q+d2)*q+d3)*q+d4)*q+1);
  }
  q = p - 0.5; r = q*q;
  return (((((a1*r+a2)*r+a3)*r+a4)*r+a5)*r+a6)*q /
         (((((b1*r+b2)*r+b3)*r+b4)*r+b5)*r+1);
}

// ===== Helpers =====
static inline double clamp_dual(double d){
  if (!std::isfinite(d)) return 0.0;
  if (d < 0.0) d = -d;
  const double DMAX = 1e6; if (d > DMAX) d = DMAX;
  return d;
}
static inline double quantize_mppc(double x, double mppc, int sw){
  if (!sw || mppc<=0.0) return x;
  double k = std::floor(x / mppc);
  return std::max(0.0, k * mppc);
}

// ===== Core solver =====
static double solve_global_policies(
  // Demand & physics
  const double* Demand,             // [R,S,H]
  const double* CF,                 // [R,E,S,H]
  const double* Hours,              // [S,H]
  const double* PeakMask,           // [S,H]

  // Region-specific generator costs
  const double* CAPEX_RE,           // [R,E]
  const double* FOM_new_RE,         // [R,E]
  const double* Fuel_RE,            // [R,E]
  const double* VOM_RE,             // [R,E]

  // Existing & caps (generation)
  const double* ExistingCap,        // [R,E]
  const double* MaxBuild,           // [R,E]
  const double* CapMax,             // [R,E]

  // Trading caps
  const double* TransRR,            // [R,R]

  // CES / RPS
  const double* CES_q, double CES_rhs, double CES_ACP, // [E]
  const double* RPS_q, double RPS_rhs, double RPS_ACP, // [E]

  // Retirement economics
  const double* FOM_exist_RE,       // [R,E] $/MW-yr existing
  const double* RetireCost_RE,      // [R,E] $/MW-yr
  const double* RevExist_RE,        // [R,E] $/MW-yr observed
  const double* RetireResponse_RE,  // [R,E] MW per ($/MW-yr)
  const double* BelowFOMFlag_RE,    // [R,E] 0/1
  double FOM_cover_mult,            // 0..1

  // Storage (4-hour only; power MW)
  const double* StorExistPower_R,   // [R] MW existing power
  const double* StorMaxBuild_R,     // [R] MW/yr
  const double* StorMaxTotal_R,     // [R] MW total
  const double* StorCapCost_R,      // [R] $/MW (power)
  const double* StorCapFOM_R,       // [R] $/MW-yr (power)
  const double* StorEnCost_R,       // [R] $/MWh (energy)
  const double* StorEnFOM_R,        // [R] $/MWh-yr (energy)
  const double* StorChgVOM_R,       // [R] $/MWh charge
  const double* StorDisVOM_R,       // [R] $/MWh discharge

  // MPPC & NSD
  const double* MPPC_E,             // [E] MW; 0 disables
  const double* NSD_E,              // [E] normalized std dev (>=0)

  // Reserve & system
  double CES_rhs_s, double CES_ACP_s,
  double RPS_rhs_s, double RPS_ACP_s,
  double FOM_cover_mult_s, double ReserveMargin,
  double StorChargeEff, double StorDischargeEff,
  double StorSoC0Frac, int StorCyclic, double StorCapCredit,
  int MinPlantSizeSwitch, int BlocksK, int AnchorMean,
  int R,int S,int H,int E
){
  const int T=S*H;
  g_S=S; g_H=H;

  // Trading enable map
  bool anyTrade=false; vector<char> allow(R*R,0);
  if (TransRR){
    for(int r=0;r<R;++r) for(int d=0; d<R; ++d){
      if(r==d) continue;
      double cap = TransRR[r*R + d];
      if (cap > 0.0){ allow[r*R + d]=1; anyTrade=true; }
    }
  }
  const bool haveTrade=anyTrade;

  // Rows used? (policy rows)
  const bool useCES=(CES_rhs>0.0)||(CES_ACP>0.0);
  const bool addCES=(CES_ACP>0.0);
  const bool useRPS=(RPS_rhs>0.0)||(RPS_ACP>0.0);
  const bool addRPS=(RPS_ACP>0.0);
  const bool useReserve=(ReserveMargin>0.0);

  // K-block setup
  int K = BlocksK<1?1:BlocksK;
  if (K>10) K=10;
  g_nsd_E.assign(E,0.0);
  for(int e0=0;e0<E;++e0) g_nsd_E[e0] = std::max(0.0, NSD_E? NSD_E[e0] : 0.0);
  g_mppc_E.assign(E,0.0);
  for(int e0=0;e0<E;++e0) g_mppc_E[e0] = MPPC_E? MPPC_E[e0] : 0.0;
  g_mppc_switch = MinPlantSizeSwitch?1:0;
  g_anchorMean  = AnchorMean?true:false;

  reset_cache(R,E,T, addCES, addRPS, haveTrade, K, g_anchorMean);

  // Save snapshots needed by getters
  g_exist_cap.assign(R*E,0.0);
  for(int r=0;r<R;++r) for(int e0=0;e0<E;++e0) g_exist_cap[r*E+e0]=ExistingCap[idx_re(r,e0,E)];
  g_hours_SH.assign(S*H,0.0);
  for(int s=0;s<S;++s) for(int h0=0;h0<H;++h0) g_hours_SH[s*H+h0]=Hours[idx_hours(s,h0,H)];
  g_cf_RESH.assign((size_t)R*E*S*H,0.0);
  for(int r=0;r<R;++r) for(int e0=0;e0<E;++e0) for(int s=0;s<S;++s) for(int h0=0;h0<H;++h0)
    g_cf_RESH[ ((r*E + e0)*S + s)*H + h0 ] = CF[idx_cf(r,e0,s,h0,E,S,H)];

  // ----- Rows -----
  // Demand R*T    | Capacity R*E*T | Reserve (opt) R*T | CES (opt) 1 | RPS (opt) 1
  // Storage rows: SoC dyn R*T | ChargeCap R*T | DisCap R*T | SoCMax R*T | SoCInit/Cycle R
  const int ROW_DEM0=0;                       // R*T
  const int ROW_CAP0=ROW_DEM0 + R*T;          // R*E*T
  int rowCount = ROW_CAP0 + R*E*T;

  // Per-block capacity rows (enforce merit-order block widths)
  const int ROW_CAPK0 = rowCount;             // R*E*K*T
  rowCount += R*E*K*T;

  const int ROW_RSV0=rowCount; if(useReserve) rowCount+=R*T;
  const int ROW_CES=rowCount;  if(useCES) rowCount+=1;
  const int ROW_RPS=rowCount;  if(useRPS) rowCount+=1;

  const int ROW_SOC_DYN0 = rowCount;          // R*T
  rowCount += R*T;
  const int ROW_ST_CHCAP0 = rowCount;         // R*T
  rowCount += R*T;
  const int ROW_ST_DCCAP0 = rowCount;         // R*T
  rowCount += R*T;
  const int ROW_ST_SOCMAX0 = rowCount;        // R*T
  rowCount += R*T;
  const int ROW_ST_SOCINIT0 = rowCount;       // R
  rowCount += R;

  // ----- Columns -----
  const double INF=1e30;
  vector<double> col_cost(g_N_total,0.0), col_lo(g_N_total,0.0), col_hi(g_N_total,INF);

  // Precompute K-block cost multipliers z_k
  vector<double> zK(K, 0.0);
  if (K==1) { zK[0]=0.0; }
  else {
    // block mid-quantiles p_k = (k+0.5)/K
    for(int k=0;k<K;++k){
      double p = (k + 0.5) / K;
      double z = inv_norm_cdf(p);
      zK[k] = z;
    }
    if (g_anchorMean) {
      // normalize z to zero-mean exactly iff requested
      double meanz=0.0; for(double z: zK) meanz+=z; meanz/=K;
      for(double& z: zK) z -= meanz;
    }
  }

  // Region-scoped columns
  for(int r=0;r<R;++r){
    const int base=r*g_colsPerReg;

    // Build gen (E)
    for(int e0=0;e0<E;++e0){
      int cB=base+col_build_e(e0);
      double capex=CAPEX_RE[idx_re(r,e0,E)];
      double fomN =FOM_new_RE[idx_re(r,e0,E)];
      col_cost[cB]=capex + fomN;
      col_lo[cB]=0.0;
      double ubB=INF;
      if(MaxBuild){ double mb=MaxBuild[idx_re(r,e0,E)]; if(mb>0.0) ubB=std::min(ubB,mb); }
      if(CapMax){ double cm=CapMax[idx_re(r,e0,E)]; if(cm>0.0){ double ex=ExistingCap[idx_re(r,e0,E)]; ubB=std::min(ubB,std::max(0.0, cm - ex)); } }
      col_hi[cB]=ubB;
    }

    // Retire gen (E)
    for(int e0=0;e0<E;++e0){
      int cR=base+col_retire_e(E,e0);
      col_lo[cR]=0.0;
      double Fexist=FOM_exist_RE     [idx_re(r,e0,E)];
      double Rexist=RevExist_RE      [idx_re(r,e0,E)];
      double slope =RetireResponse_RE[idx_re(r,e0,E)];
      double flag  =BelowFOMFlag_RE  [idx_re(r,e0,E)];
      double need  =FOM_cover_mult * Fexist;
      double shortfall = std::max(0.0, need - Rexist);
      double econCap = (flag>0.5 && shortfall>0.0 && slope>0.0) ? shortfall * slope : 0.0;
      double ubR = std::min( ExistingCap[idx_re(r,e0,E)], econCap );
      if(ubR<0.0) ubR=0.0;
      col_hi[cR]=ubR;
      double retireCost=RetireCost_RE[idx_re(r,e0,E)];
      col_cost[cR]=retireCost - Fexist;
    }

    // Generation per block
    for(int e0=0;e0<E;++e0){
      const double base_v = Fuel_RE[idx_re(r,e0,E)] + VOM_RE[idx_re(r,e0,E)];
      const double nsd    = g_nsd_E[e0];
      for(int k=0;k<K;++k){
        const double mult = 1.0 + nsd * zK[k];
        const double varc = base_v * std::max(0.0, mult);
        for(int tau=0;tau<T;++tau){
          int cG = base + col_gen_e_k_tau(E,K,e0,k,tau,T);
          col_cost[cG]=varc; col_lo[cG]=0.0; col_hi[cG]=INF;
        }
      }
    }

    // Flows
    if(haveTrade){
      for(int d=0; d<R; ++d){
        if(d==r) continue;
        bool ok = allow[r*R + d];
        for(int tau=0;tau<T;++tau){
          int cF=base+col_flow_od_tau(E,K,T,R, r,d,tau);
          col_cost[cF]=0.0; col_lo[cF]=0.0; col_hi[cF]= ok? TransRR[r*R + d] : 0.0;
        }
      }
    }

    // Storage build power (MW)
    {
      int cSB = base + col_stor_build(E,K,T,R,haveTrade);
      col_lo[cSB]=0.0;
      double ub=INF;
      if (StorMaxBuild_R){ double mb=StorMaxBuild_R[r]; if(mb>0.0) ub=std::min(ub, mb); }
      if (StorMaxTotal_R){ double mt=StorMaxTotal_R[r]; if(mt>0.0){ double exP=StorExistPower_R[r]; ub=std::min(ub, std::max(0.0, mt - exP)); } }
      col_hi[cSB]=ub;

      // cost: power capex+FOM + 4*(energy capex+FOM) per MW of power (4-hr)
      double cap_cost = StorCapCost_R?  StorCapCost_R[r]  : 0.0;
      double cap_fom  = StorCapFOM_R?   StorCapFOM_R[r]   : 0.0;
      double en_cost  = StorEnCost_R?   StorEnCost_R[r]   : 0.0;
      double en_fom   = StorEnFOM_R?    StorEnFOM_R[r]    : 0.0;
      col_cost[cSB] = (cap_cost + cap_fom) + 4.0*(en_cost + en_fom);
    }

    // Storage charge/discharge/SoC bounds
    for(int tau=0;tau<T;++tau){
      int cC = base + col_stor_charge_tau(E,K,T,R,haveTrade,tau);
      int cD = base + col_stor_discharge_tau(E,K,T,R,haveTrade,tau);
      int cS = base + col_stor_soc_tau(E,K,T,R,haveTrade,tau);
      col_lo[cC]=0.0; col_hi[cC]=INF; col_cost[cC] = StorChgVOM_R? StorChgVOM_R[r] : 0.0;
      col_lo[cD]=0.0; col_hi[cD]=INF; col_cost[cD] = StorDisVOM_R? StorDisVOM_R[r] : 0.0;
      col_lo[cS]=0.0; col_hi[cS]=INF; // SoC bounded by rows
    }
  }

  // Slacks
  if(g_colSlackCES>=0){ col_cost[g_colSlackCES]=CES_ACP; col_lo[g_colSlackCES]=0.0; col_hi[g_colSlackCES]=INF; }
  if(g_colSlackRPS>=0){ col_cost[g_colSlackRPS]=RPS_ACP; col_lo[g_colSlackRPS]=0.0; col_hi[g_colSlackRPS]=INF; }

  // ----- Row bounds -----
  vector<double> row_lo(rowCount,-INF), row_hi(rowCount,INF);

  auto is_peak = [&](int s,int h0)->bool{
    if(!PeakMask) return true;
    return PeakMask[idx_mask(s,h0,H)] > 0.5;
  };

  // Demand (=)
  for(int r=0;r<R;++r)
    for(int s=0;s<S;++s)
      for(int h0=0;h0<H;++h0){
        int tau=s*H+h0, row=ROW_DEM0 + r*T + tau;
        double rhs=Demand[idx_dem(r,s,h0,S,H)];
        row_lo[row]=rhs; row_hi[row]=rhs;
      }

  // Capacity (>=) per (r,e,tau): -Σk Gen + CF*hrs*(Build - Retire) >= -CF*hrs*Existing
  for(int r=0;r<R;++r)
    for(int e0=0;e0<E;++e0){
      double ex=ExistingCap[idx_re(r,e0,E)];
      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0, row=ROW_CAP0 + ((r*E+e0)*T + tau);
          double cf = CF[idx_cf(r,e0,s,h0,E,S,H)];
          double hrs= Hours[idx_hours(s,h0,H)];
          row_lo[row] = -cf*hrs*ex; row_hi[row] = INF;
        }
    }

  // Reserve rows (>=) only at peak slices.
  if(useReserve){
    for(int r=0;r<R;++r)
      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0, row=ROW_RSV0 + r*T + tau;
          if(!is_peak(s,h0)){ row_lo[row]=-INF; row_hi[row]=INF; continue; }
          double hrs=Hours[idx_hours(s,h0,H)];
          double rhs=(1.0+ReserveMargin)*Demand[idx_dem(r,s,h0,S,H)];
          for(int e0=0;e0<E;++e0) rhs -= CF[idx_cf(r,e0,s,h0,E,S,H)] * hrs * ExistingCap[idx_re(r,e0,E)];
          // subtract storage existing (credit*hrs*P_exist)
          double Pexist = StorExistPower_R? StorExistPower_R[r] : 0.0;
          rhs -= StorCapCredit * hrs * Pexist;
          row_lo[row]=rhs; row_hi[row]=INF;
        }
  }

  // CES / RPS (>=)
  if(useCES){ row_lo[ROW_CES]=std::max(0.0,CES_rhs); row_hi[ROW_CES]=INF; }
  if(useRPS){ row_lo[ROW_RPS]=std::max(0.0,RPS_rhs); row_hi[ROW_RPS]=INF; }

  // Per-block capacity bounds with equal widths alpha_k = 1/K
  {
    double alpha_k = (K>0) ? (1.0 / K) : 1.0;
    for(int r=0;r<R;++r)
      for(int e0=0;e0<E;++e0)
        for(int k=0;k<K;++k)
          for(int s=0;s<S;++s)
            for(int h0=0;h0<H;++h0){
              int tau=s*H+h0;
              int row = ROW_CAPK0 + (((r*E + e0)*K + k)*T + tau);
              double cf = CF[idx_cf(r,e0,s,h0,E,S,H)];
              double hrs= Hours[idx_hours(s,h0,H)];
              double ex = ExistingCap[idx_re(r,e0,E)];
              row_lo[row] = -alpha_k * cf * hrs * ex;
              row_hi[row] = INF;
            }
  }

  // Storage rows
  for(int r=0;r<R;++r){
    double Pexist = StorExistPower_R? StorExistPower_R[r] : 0.0;
    double Eexist = 4.0 * Pexist;
    // SoC dynamics: SoC[t] - SoC[t-1] - ηc*Chg + (1/η_d)*Dis = 0
    for(int tau=0;tau<T;++tau){
      int row = ROW_SOC_DYN0 + r*T + tau;
      row_lo[row]=0.0; row_hi[row]=0.0; // equality
    }
    // ChargeCap: Chg[t] - BuildPower <= Pexist
    for(int tau=0;tau<T;++tau){
      int row = ROW_ST_CHCAP0 + r*T + tau;
      row_lo[row]=-INF; row_hi[row]=Pexist;
    }
    // DisCap: Dis[t] - BuildPower <= Pexist
    for(int tau=0;tau<T;++tau){
      int row = ROW_ST_DCCAP0 + r*T + tau;
      row_lo[row]=-INF; row_hi[row]=Pexist;
    }
    // SoCMax: SoC[t] - 4*BuildPower <= 4*Pexist
    for(int tau=0;tau<T;++tau){
      int row = ROW_ST_SOCMAX0 + r*T + tau;
      row_lo[row]=-INF; row_hi[row]=Eexist;
    }
    // SoC init/cycle:
    int rowI = ROW_ST_SOCINIT0 + r;
    if (StorCyclic){
      row_lo[rowI]=0.0; row_hi[rowI]=0.0; // SoC[T-1] - SoC[0] = 0
    } else {
      row_lo[rowI]=Eexist * StorSoC0Frac;  // SoC[0] - 4*SoC0*Build = Eexist*SoC0
      row_hi[rowI]=Eexist * StorSoC0Frac;
    }
  }

  // ----- Triplets -----
  vector<Triplet> Tpls;
  // Rough reserve:
  Tpls.reserve((size_t)R*E*T*(2 + g_K) + (useReserve? (size_t)R*T*(E+ (g_haveTrade?R:0) + 1):0)
               + (useCES? (size_t)R*E*T + (g_colSlackCES>=0?1:0) : 0)
               + (useRPS? (size_t)R*E*T + (g_colSlackRPS>=0?1:0) : 0)
               + (size_t)R*(g_haveTrade?(R-1):0)*T*2
               + (size_t)R*(T*6 + 1) // storage rows coeff estimates
               + (size_t)R*E*g_K*T   // per-block capacity rows
               );

  for(int r=0;r<R;++r){
    int base=r*g_colsPerReg;

    for(int s=0;s<S;++s)
      for(int h0=0;h0<H;++h0){
        int tau=s*H+h0; double hrs=Hours[idx_hours(s,h0,H)];

        // Demand: Σe,k Gen + inflows − outflows + Dis − Ch = Demand
        int rowD=ROW_DEM0 + r*T + tau;
        for(int e0=0;e0<E;++e0)
          for(int k=0;k<K;++k)
            Tpls.push_back({rowD, base+col_gen_e_k_tau(E,K,e0,k,tau,T), 1.0});

        if(g_haveTrade){
          for(int o=0;o<R;++o){
            if(o==r) continue;
            if(!(TransRR && TransRR[o*R + r] > 0.0)) continue;
            int base_o=o*g_colsPerReg;
            Tpls.push_back({rowD, base_o+col_flow_od_tau(E,K,T,R, o,r,tau), 1.0});
          }
          for(int d=0; d<R; ++d){
            if(d==r) continue;
            if(!(TransRR && TransRR[r*R + d] > 0.0)) continue;
            Tpls.push_back({rowD, base+col_flow_od_tau(E,K,T,R, r,d,tau), -1.0});
          }
        }
        // + Discharge - Charge
        Tpls.push_back({rowD, base+col_stor_discharge_tau(E,K,T,R,g_haveTrade,tau), 1.0});
        Tpls.push_back({rowD, base+col_stor_charge_tau   (E,K,T,R,g_haveTrade,tau), -1.0});

        // Capacity rows: -Σk Gen + CF*hrs*(Build - Retire)
        for(int e0=0;e0<E;++e0){
          double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
          int rowC=ROW_CAP0 + ((r*E+e0)*T + tau);
          for(int k=0;k<K;++k)
            Tpls.push_back({rowC, base+col_gen_e_k_tau(E,K,e0,k,tau,T), -1.0});
          if(cf!=0.0 && hrs!=0.0){
            Tpls.push_back({rowC, base+col_build_e(e0),  cf*hrs});
            Tpls.push_back({rowC, base+col_retire_e(E,e0), -cf*hrs});
          }
        }

        // Per-block capacity rows (equal widths): -Gen_k + alpha_k*CF*hrs*(Build - Retire) >= -alpha_k*CF*hrs*Existing
        {
          double alpha_k = (K>0) ? (1.0 / K) : 1.0;
          for(int e0=0;e0<E;++e0){
            double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
            if(cf!=0.0 && hrs!=0.0){
              for(int k=0;k<K;++k){
                int rowCK = ROW_CAPK0 + (((r*E + e0)*K + k)*T + tau);
                // -Gen_k
                Tpls.push_back({rowCK, base+col_gen_e_k_tau(E,K,e0,k,tau,T), -1.0});
                // + alpha_k * CF * hrs * Build
                Tpls.push_back({rowCK, base+col_build_e(e0),  alpha_k * cf * hrs});
                // - alpha_k * CF * hrs * Retire
                Tpls.push_back({rowCK, base+col_retire_e(E,e0), -alpha_k * cf * hrs});
              }
            }
          }
        }

        // Reserve rows (peak): gen build-retire contributes; plus imports; plus storage credit*hrs*BuildPower
        if(useReserve && is_peak(s,h0)){
          int rowR=ROW_RSV0 + r*T + tau;
          for(int e0=0;e0<E;++e0){
            double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
            if(cf!=0.0 && hrs!=0.0){
              Tpls.push_back({rowR, base+col_build_e(e0),  cf*hrs});
              Tpls.push_back({rowR, base+col_retire_e(E,e0), -cf*hrs});
            }
          }
          if(g_haveTrade){
            for(int o=0;o<R;++o){
              if(o==r) continue;
              if(!(TransRR && TransRR[o*R + r] > 0.0)) continue;
              int base_o=o*g_colsPerReg;
              Tpls.push_back({rowR, base_o+col_flow_od_tau(E,K,T,R, o,r,tau), 1.0});
            }
          }
          // storage contribution: + credit * hrs * BuildPower
          Tpls.push_back({rowR, base+col_stor_build(E,K,T,R,g_haveTrade), StorCapCredit * hrs});
        }

        // CES / RPS rows
        if(useCES){
          int rowC=ROW_CES;
          for(int e0=0;e0<E;++e0){ double q=CES_q[e0]; if(q!=0.0)
            for(int k=0;k<K;++k) Tpls.push_back({rowC, base+col_gen_e_k_tau(E,K,e0,k,tau,T), q});
          }
        }
        if(useRPS){
          int rowR=ROW_RPS;
          for(int e0=0;e0<E;++e0){ double q=RPS_q[e0]; if(q!=0.0)
            for(int k=0;k<K;++k) Tpls.push_back({rowR, base+col_gen_e_k_tau(E,K,e0,k,tau,T), q});
          }
        }

        // Storage rows for this tau
        int cSB = base+col_stor_build(E,K,T,R,g_haveTrade);
        int cC  = base+col_stor_charge_tau(E,K,T,R,g_haveTrade,tau);
        int cD  = base+col_stor_discharge_tau(E,K,T,R,g_haveTrade,tau);
        int cS  = base+col_stor_soc_tau(E,K,T,R,g_haveTrade,tau);
        int cS_prev = base+col_stor_soc_tau(E,K,T,R,g_haveTrade, (tau==0? T-1 : tau-1));

        // SoC dyn: SoC[t] - SoC[t-1] - eta_c*Chg + (1/eta_d)*Dis = 0
        int rowSD = ROW_SOC_DYN0 + r*T + tau;
        Tpls.push_back({rowSD, cS,  1.0});
        Tpls.push_back({rowSD, cS_prev, -1.0});
        if (StorChargeEff!=0.0) Tpls.push_back({rowSD, cC, -StorChargeEff});
        if (StorDischargeEff!=0.0) Tpls.push_back({rowSD, cD,  1.0/StorDischargeEff});

        // ChargeCap: Chg[t] - BuildPower <= Pexist
        int rowCC = ROW_ST_CHCAP0 + r*T + tau;
        Tpls.push_back({rowCC, cC, 1.0});
        Tpls.push_back({rowCC, cSB, -1.0});

        // DisCap: Dis[t] - BuildPower <= Pexist
        int rowDC = ROW_ST_DCCAP0 + r*T + tau;
        Tpls.push_back({rowDC, cD, 1.0});
        Tpls.push_back({rowDC, cSB, -1.0});

        // SoCMax: SoC[t] - 4*BuildPower <= 4*Pexist
        int rowSM = ROW_ST_SOCMAX0 + r*T + tau;
        Tpls.push_back({rowSM, cS, 1.0});
        Tpls.push_back({rowSM, cSB, -4.0});
      }

    // SoC init or cycle
    int cS0 = base+col_stor_soc_tau(E,K,T,R,g_haveTrade,0);
    int cSLast = base+col_stor_soc_tau(E,K,T,R,g_haveTrade,T-1);
    int cSB   = base+col_stor_build(E,K,T,R,g_haveTrade);
    int rowI = ROW_ST_SOCINIT0 + r;
    if (StorCyclic){
      // SoC[T-1] - SoC[0] = 0
      Tpls.push_back({rowI, cSLast,  1.0});
      Tpls.push_back({rowI, cS0,    -1.0});
    } else {
      // SoC[0] - 4*SoC0*BuildPower = 4*SoC0*Eexist
      Tpls.push_back({rowI, cS0, 1.0});
      Tpls.push_back({rowI, cSB, -4.0*StorSoC0Frac});
    }
  }

  // Slacks
  if(useCES && g_colSlackCES>=0) Tpls.push_back({ROW_CES,g_colSlackCES,1.0});
  if(useRPS && g_colSlackRPS>=0) Tpls.push_back({ROW_RPS,g_colSlackRPS,1.0});

  // Convert & solve
  vector<int> astart_i,aindex_i; vector<double> avalue;
  triplets_to_csc(rowCount,g_N_total,Tpls,astart_i,aindex_i,avalue);
  vector<HighsInt> astart(astart_i.begin(),astart_i.end()), aindex(aindex_i.begin(),aindex_i.end());
  HighsInt nnz=(HighsInt)avalue.size();
  void* h=Highs_create();
  const HighsInt a_format=1, sense=1; const double obj_offset=0.0;

  g_lastCode = Highs_passLp(h,(HighsInt)g_N_total,(HighsInt)rowCount,nnz, a_format,sense,obj_offset,
                            g_N_total? &col_cost[0]:nullptr, g_N_total? &col_lo[0]:nullptr, g_N_total? &col_hi[0]:nullptr,
                            rowCount? &row_lo[0]:nullptr, rowCount? &row_hi[0]:nullptr,
                            astart.data(), aindex.data(), avalue.data());
  if(g_lastCode==0) g_lastCode=Highs_run(h);

  g_status=(g_lastCode==0)? Highs_getModelStatus(h) : -1;
  g_obj=1e308; g_solved=false;

  if(g_lastCode==0){
    double obj=Highs_getObjectiveValue(h);
    vector<double> col_val(g_N_total,0.0), col_dual(g_N_total,0.0);
    vector<double> row_val(std::max(1,rowCount),0.0), row_dual(std::max(1,rowCount),0.0);
    if(Highs_getSolution(h,col_val.data(),col_dual.data(),row_val.data(),row_dual.data())==0){
      g_solution=std::move(col_val);

      // LMPs from demand duals
      g_lmp.assign(R*T,0.0);
      for(int r=0;r<R;++r) for(int tau=0;tau<T;++tau){
        int row=ROW_DEM0 + r*T + tau;
        g_lmp[r*T + tau] = row_dual[row];
      }

      // Reserve duals — peak-only, clamp
      g_capdual.assign(R*T, 0.0);
      if (useReserve){
        for(int r=0;r<R;++r) for(int s=0;s<S;++s) for(int h0=0;h0<H;++h0){
          int tau=s*H + h0, row=ROW_RSV0 + r*T + tau;
          double d = is_peak(s,h0) ? clamp_dual(row_dual[row]) : 0.0;
          g_capdual[r*T + tau] = d;
        }
      }

      // Energy revenue per (r,e)
      g_rev_tech.assign(R*E,0.0);
      for(int r=0;r<R;++r){
        int base=r*g_colsPerReg;
        for(int e0=0;e0<E;++e0){
          double rev=0.0;
          for(int tau=0;tau<T;++tau){
            // sum across K blocks
            double gen_sum = 0.0;
            for(int k=0;k<K;++k){
              int cG=base+col_gen_e_k_tau(E,K,e0,k,tau,T);
              gen_sum += g_solution[cG];
            }
            rev += g_lmp[r*T + tau] * gen_sum;
          }
          g_rev_tech[r*E + e0]=rev;
        }
      }

      // Regional cost breakout
      for(int r=0;r<R;++r){
        double cc=0.0, ce=0.0; int base=r*g_colsPerReg;
        for(int e0=0;e0<E;++e0){
          int cB=base+col_build_e(e0), cR=base+col_retire_e(E,e0);
          double capex = CAPEX_RE    [idx_re(r,e0,E)];
          double fomN  = FOM_new_RE  [idx_re(r,e0,E)];
          double retire= RetireCost_RE[idx_re(r,e0,E)];
          double Fexist= FOM_exist_RE[idx_re(r,e0,E)];
          cc += (capex + fomN) * g_solution[cB] + (retire - Fexist) * g_solution[cR];
          // variable energy
          for(int tau=0;tau<T;++tau){
            for(int k=0;k<K;++k){
              int cG=base+col_gen_e_k_tau(E,K,e0,k,tau,T);
              double base_v = Fuel_RE[idx_re(r,e0,E)] + VOM_RE[idx_re(r,e0,E)];
              double nsd    = g_nsd_E[e0];
              double mult   = 1.0 + nsd * 0.0; // energy cost accounting at base variable cost
              (void)mult; // keep energy cost consistent with objective? choose base fuel+VOM:
              ce += (Fuel_RE[idx_re(r,e0,E)] + VOM_RE[idx_re(r,e0,E)]) * g_solution[cG];
            }
          }
        }
        // storage variable costs
        int cSB = base+col_stor_build(E,K,T,R,g_haveTrade);
        double cap_cost = StorCapCost_R?  StorCapCost_R[r]  : 0.0;
        double cap_fom  = StorCapFOM_R?   StorCapFOM_R[r]   : 0.0;
        double en_cost  = StorEnCost_R?   StorEnCost_R[r]   : 0.0;
        double en_fom   = StorEnFOM_R?    StorEnFOM_R[r]    : 0.0;
        cc += ((cap_cost + cap_fom) + 4.0*(en_cost + en_fom)) * g_solution[cSB];
        for(int tau=0;tau<T;++tau){
          ce += (StorChgVOM_R? StorChgVOM_R[r]:0.0) * g_solution[base+col_stor_charge_tau(E,K,T,R,g_haveTrade,tau)];
          ce += (StorDisVOM_R? StorDisVOM_R[r]:0.0) * g_solution[base+col_stor_discharge_tau(E,K,T,R,g_haveTrade,tau)];
        }

        g_cost_cap[r]=cc; g_cost_energy[r]=ce;
      }

      g_obj=obj; g_solved=std::isfinite(g_obj);
    }
  }

  Highs_destroy(h);
  return g_solved? g_obj : 1e308;
}

// ===== Vensim glue =====
static const int EXTERN_VCODE=62051;
extern "C" __declspec(dllexport) int VEFCC version_info(){ return EXTERN_VCODE; }

enum {
  F_SOLVE=1101,
  F_CAP_ADD, F_RETIRE, F_CAP_INST, F_GEN,
  F_COST_CAP, F_COST_ENERGY, F_LP_OBJ, F_LP_STATUS, F_LP_CODE,
  F_FLOW_EXPORT, F_FLOW_IMPORT, F_FLOW_OD,
  F_LP_PRICE, F_LP_CAPSLICE_PRICE, F_LP_REV_TECH,
  F_LP_REV_EN_PERMW_INST, F_LP_REV_CAP_PERMW_INST,
  F_LP_CAP_PRICE_TECH, F_LP_CAP_PRICE_REGION,
  // Storage
  F_STOR_BUILD_CAP, F_STOR_CHARGE, F_STOR_DISCHARGE, F_STOR_SOC
};

extern "C" __declspec(dllexport) int VEFCC user_definition(
  int setup_index, char** sym, char** arglist,
  int* num_args, int* num_vector, int* func_index,
  int* dim_act, int* modify, int* num_loops,
  int* num_literal, int* num_lookup)
{
  if(!sym||!arglist||!num_args||!num_vector||!func_index||!dim_act||!modify||!num_loops||!num_literal||!num_lookup) return 0;
  *dim_act=0; *modify=0; *num_loops=0; *num_literal=0; *num_lookup=0;

  switch(setup_index){
    case 0:
      *sym=(char*)"LP_Solve";
      // VECTORS (30): Demand, CF, Hours, PeakMask,
      // CAPEX,FOM_new,Fuel,VOM, ExistingCap,MaxBuild,CapMax,TransRR, CES_q,RPS_q,
      // FOM_exist,RetireCost,RevExist,RetireResponse,BelowFOMFlag,
      // StorExistP,StorMaxBuild,StorMaxTotal,StorCapCost,StorCapFOM,StorEnCost,StorEnFOM,StorChgVOM,StorDisVOM,
      // MPPC_E, NSD_E
      // SCALARS (19): CES_rhs,CES_ACP,RPS_rhs,RPS_ACP,FOM_cover_mult,ReserveMargin,
      // StorChargeEff,StorDischargeEff,StorSoC0Frac,StorCyclic,StorCapCredit,
      // MinPlantSizeSwitch,BlocksK,AnchorMean, R,S,H,E, scenario_id
      *arglist=(char*)"{Demand},{Capacity_factor},{Hours},{PeakMask},{CAPEX},{FOM_new},{Fuel},{VOM},{ExistingCap},{MaxBuild},{CapMax},{TransRR},{CES_qualifying},{RPS_qualifying},{FOM_exist},{RetireCost},{RevExist},{RetireResponse},{BelowFOMFlag},{StorExistPower},{StorMaxBuild},{StorMaxTotal},{StorCapCost},{StorCapFOM},{StorEnCost},{StorEnFOM},{StorChgVOM},{StorDisVOM},{MPPC_Min},{NSD},{CES_rhs},{CES_ACP},{RPS_rhs},{RPS_ACP},{FOM_cover_mult},{ReserveMargin},{StorChargeEff},{StorDischargeEff},{StorSoC0Frac},{StorCyclic},{StorCapCredit},{MinPlantSizeSwitch},{BlocksK},{AnchorMean},R,S,H,E,scenario_id";
      *num_args=49; *num_vector=30; *func_index=F_SOLVE; return 1;

    case 1:  *sym=(char*)"LP_CAP_ADD";       *arglist=(char*)" region_pos , tech_pos , scenario_id ";                  *num_args=3; *num_vector=0; *func_index=F_CAP_ADD;      return 1;
    case 2:  *sym=(char*)"LP_RETIRE";        *arglist=(char*)" region_pos , tech_pos , scenario_id ";                  *num_args=3; *num_vector=0; *func_index=F_RETIRE;       return 1;
    case 3:  *sym=(char*)"LP_CAP_INSTALLED"; *arglist=(char*)" region_pos , tech_pos , scenario_id ";                  *num_args=3; *num_vector=0; *func_index=F_CAP_INST;     return 1;
    case 4:  *sym=(char*)"LP_GEN";           *arglist=(char*)" region_pos , tech_pos , ts_pos , hr_pos , scenario_id "; *num_args=5; *num_vector=0; *func_index=F_GEN;          return 1;

    case 5:  *sym=(char*)"LP_COST_CAP";      *arglist=(char*)" region_pos , scenario_id ";                              *num_args=2; *num_vector=0; *func_index=F_COST_CAP;     return 1;
    case 6:  *sym=(char*)"LP_COST_ENERGY";   *arglist=(char*)" region_pos , scenario_id ";                              *num_args=2; *num_vector=0; *func_index=F_COST_ENERGY;  return 1;
    case 7:  *sym=(char*)"LP_OBJ";           *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_OBJ;       return 1;
    case 8:  *sym=(char*)"LP_STATUS";        *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_STATUS;    return 1;
    case 9:  *sym=(char*)"LP_CODE";          *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_CODE;      return 1;

    case 10: *sym=(char*)"LP_EXPORT";        *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_FLOW_EXPORT;  return 1;
    case 11: *sym=(char*)"LP_IMPORT";        *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_FLOW_IMPORT;  return 1;
    case 12: *sym=(char*)"LP_FLOW_OD";       *arglist=(char*)" origin_pos , dest_pos , ts_pos , hr_pos , scenario_id ";*num_args=5; *num_vector=0; *func_index=F_FLOW_OD;     return 1;

    case 13: *sym=(char*)"LP_PRICE";         *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_LP_PRICE;     return 1;
    case 14: *sym=(char*)"LP_CAPSLICE_PRICE";*arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_LP_CAPSLICE_PRICE; return 1;
    case 15: *sym=(char*)"LP_REV_TECH";      *arglist=(char*)" region_pos , tech_pos , scenario_id ";                  *num_args=3; *num_vector=0; *func_index=F_LP_REV_TECH;  return 1;

    case 16: *sym=(char*)"LP_REV_EN_PERMW_INSTALLED";  *arglist=(char*)" region_pos , tech_pos , scenario_id ";       *num_args=3; *num_vector=0; *func_index=F_LP_REV_EN_PERMW_INST;  return 1;
    case 17: *sym=(char*)"LP_REV_CAP_PERMW_INSTALLED"; *arglist=(char*)" region_pos , tech_pos , scenario_id ";       *num_args=3; *num_vector=0; *func_index=F_LP_REV_CAP_PERMW_INST; return 1;

    case 18: *sym=(char*)"LP_CAP_PRICE_TECH";   *arglist=(char*)" region_pos , tech_pos , scenario_id ";              *num_args=3; *num_vector=0; *func_index=F_LP_CAP_PRICE_TECH;   return 1;
    case 19: *sym=(char*)"LP_CAP_PRICE_REGION"; *arglist=(char*)" region_pos , scenario_id ";                         *num_args=2; *num_vector=0; *func_index=F_LP_CAP_PRICE_REGION; return 1;

    // Storage getters
    case 20: *sym=(char*)"LP_STOR_BUILD_CAP"; *arglist=(char*)" region_pos , scenario_id ";                           *num_args=2; *num_vector=0; *func_index=F_STOR_BUILD_CAP; return 1;
    case 21: *sym=(char*)"LP_STOR_CHARGE";    *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";         *num_args=4; *num_vector=0; *func_index=F_STOR_CHARGE;    return 1;
    case 22: *sym=(char*)"LP_STOR_DISCHARGE"; *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";         *num_args=4; *num_vector=0; *func_index=F_STOR_DISCHARGE; return 1;
    case 23: *sym=(char*)"LP_STOR_SOC";       *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";         *num_args=4; *num_vector=0; *func_index=F_STOR_SOC;       return 1;

    default: return 0;
  }
}

extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch(funcid){

    case F_SOLVE: {
      if(nval<49){ val[0].val=1e308; return 0; }

      const double* Dem     = val[0].vec->firstval;   // [R,S,H]
      const double* CF      = val[1].vec->firstval;   // [R,E,S,H]
      const double* Hours   = val[2].vec->firstval;   // [S,H]
      const double* PeakM   = val[3].vec->firstval;   // [S,H]

      const double* CAPEX   = val[4].vec->firstval;   // [R,E]
      const double* FOMnew  = val[5].vec->firstval;   // [R,E]
      const double* Fuel    = val[6].vec->firstval;   // [R,E]
      const double* VOM     = val[7].vec->firstval;   // [R,E]

      const double* XCap    = val[8].vec->firstval;   // [R,E]
      const double* MaxB    = val[9].vec->firstval;   // [R,E]
      const double* CapMx   = val[10].vec->firstval;  // [R,E]
      const double* Trans   = val[11].vec->firstval;  // [R,R]

      const double* CES_q   = val[12].vec->firstval;  // [E]
      const double* RPS_q   = val[13].vec->firstval;  // [E]

      const double* FOMexist= val[14].vec->firstval;  // [R,E]
      const double* RetCost = val[15].vec->firstval;  // [R,E]
      const double* RevExist= val[16].vec->firstval;  // [R,E]
      const double* RetResp = val[17].vec->firstval;  // [R,E]
      const double* FlagBF  = val[18].vec->firstval;  // [R,E] 0/1

      const double* StorExistP = val[19].vec->firstval; // [R]
      const double* StorMaxB   = val[20].vec->firstval; // [R]
      const double* StorMaxTot = val[21].vec->firstval; // [R]
      const double* StorCapC   = val[22].vec->firstval; // [R]
      const double* StorCapF   = val[23].vec->firstval; // [R]
      const double* StorEnC    = val[24].vec->firstval; // [R]
      const double* StorEnF    = val[25].vec->firstval; // [R]
      const double* StorChgV   = val[26].vec->firstval; // [R]
      const double* StorDisV   = val[27].vec->firstval; // [R]

      const double* MPPC_E     = val[28].vec->firstval; // [E]
      const double* NSD_E      = val[29].vec->firstval; // [E]

      double CES_rhs       = val[30].val;
      double CES_ACP       = val[31].val;
      double RPS_rhs       = val[32].val;
      double RPS_ACP       = val[33].val;
      double FOM_cover_mult= val[34].val;
      double ReserveMargin = val[35].val;

      double StorChargeEff     = val[36].val;
      double StorDischargeEff  = val[37].val;
      double StorSoC0Frac      = val[38].val;
      int    StorCyclic        = (int)(val[39].val + 0.5);
      double StorCapCredit     = val[40].val;

      int MinPlantSizeSwitch   = (int)(val[41].val + 0.5);
      int BlocksK              = (int)(val[42].val + 0.5);
      int AnchorMean           = (int)(val[43].val + 0.5);

      int R=(int)(val[44].val+0.5), S=(int)(val[45].val+0.5), H=(int)(val[46].val+0.5), E=(int)(val[47].val+0.5);
      int scenario_id = (int)(val[48].val+0.5);

      // Persist arrays used by getters
      g_hours_SH.assign(S*H,0.0);
      for(int s=0;s<S;++s) for(int h=0;h<H;++h) g_hours_SH[s*H+h]=Hours[s*H+h];
      g_cf_RESH.assign((size_t)R*E*S*H,0.0);
      for(int r=0;r<R;++r) for(int e=0;e<E;++e) for(int s=0;s<S;++s) for(int h=0;h<H;++h)
        g_cf_RESH[ ((r*E + e)*S + s)*H + h ] = CF[ ((r*E + e)*S + s)*H + h ];
      g_exist_cap.assign(R*E,0.0);
      for(int r=0;r<R;++r) for(int e=0;e<E;++e) g_exist_cap[r*E+e]=XCap[r*E+e];

      val[0].val = solve_global_policies(
        Dem,CF,Hours,PeakM,
        CAPEX,FOMnew,Fuel,VOM,
        XCap,MaxB,CapMx,
        Trans,
        CES_q,CES_rhs,CES_ACP, RPS_q,RPS_rhs,RPS_ACP,
        FOMexist,RetCost,RevExist,RetResp,FlagBF,
        FOM_cover_mult,
        StorExistP,StorMaxB,StorMaxTot,StorCapC,StorCapF,StorEnC,StorEnF,StorChgV,StorDisV,
        MPPC_E,NSD_E,
        CES_rhs,CES_ACP,RPS_rhs,RPS_ACP,FOM_cover_mult,ReserveMargin,
        StorChargeEff,StorDischargeEff,StorSoC0Frac,StorCyclic,StorCapCredit,
        MinPlantSizeSwitch,BlocksK,AnchorMean,
        R,S,H,E
      );
      // Save context, plus store arrays needed for reserve RHS later (no renames, but we don't need extra now)
      save_context(scenario_id);
      return 1;
    }

    // --- Getters with MPPC fast quantize for build/retire ---
    case F_CAP_ADD: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; double raw = c->solution[base+col_build_e(e-1)];
      double mppc = (e-1)<(int)c->mppc_E.size()? c->mppc_E[e-1] : 0.0;
      val[0].val = quantize_mppc(raw, mppc, c->mppc_switch);
      return 1;
    }

    case F_RETIRE: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; double raw = c->solution[base+col_retire_e(c->E,e-1)];
      double mppc = (e-1)<(int)c->mppc_E.size()? c->mppc_E[e-1] : 0.0;
      val[0].val = quantize_mppc(raw, mppc, c->mppc_switch);
      return 1;
    }

    // ΔInstalled = quantized(Build) − quantized(Retire)
    case F_CAP_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg;
      double b = c->solution[base+col_build_e(e-1)];
      double q = c->solution[base+col_retire_e(c->E,e-1)];
      double mppc = (e-1)<(int)c->mppc_E.size()? c->mppc_E[e-1] : 0.0;
      b = quantize_mppc(b, mppc, c->mppc_switch);
      q = quantize_mppc(q, mppc, c->mppc_switch);
      val[0].val = b - q;
      return 1;
    }

    case F_GEN: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; int tau=(ts-1)*c->H+(hr-1);
      // Sum across K blocks
      double sum=0.0;
      for(int k=0;k<c->K;++k) sum += c->solution[base+col_gen_e_k_tau(c->E,c->K,e-1,k,tau,c->T)];
      val[0].val=sum; return 1;
    }

    case F_COST_CAP:    { if(nval<2){ val[0].val=1e308; return 1; } int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5); const Context* c=get_ctx(sid); if(!c||!c->solved||r<1||r>c->R){ val[0].val=1e308; return 1; } val[0].val=c->cost_cap[r-1];    return 1; }
    case F_COST_ENERGY: { if(nval<2){ val[0].val=1e308; return 1; } int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5); const Context* c=get_ctx(sid); if(!c||!c->solved||r<1||r>c->R){ val[0].val=1e308; return 1; } val[0].val=c->cost_energy[r-1]; return 1; }
    case F_LP_OBJ:      { if(nval<1){ val[0].val=1e308; return 1; } int sid=(int)(val[0].val+0.5); const Context* c=get_ctx(sid); val[0].val=c?c->obj:1e308;    return 1; }
    case F_LP_STATUS:   { if(nval<1){ val[0].val=1e308; return 1; } int sid=(int)(val[0].val+0.5); const Context* c=get_ctx(sid); val[0].val=(double)(c?c->status:-1); return 1; }
    case F_LP_CODE:     { if(nval<1){ val[0].val=1e308; return 1; } int sid=(int)(val[0].val+0.5); const Context* c=get_ctx(sid); val[0].val=(double)(c?c->lastCode:-999); return 1; }

    // Flows
    case F_FLOW_EXPORT: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      if(!c->haveTrade){ val[0].val=0.0; return 1; }
      int r0=r-1, tau=(ts-1)*c->H+(hr-1), base_r=r0*c->colsPerReg;
      double sum=0.0;
      for(int d0=0; d0<c->R; ++d0){ if(d0==r0) continue; int col=base_r+col_flow_od_tau(c->E,c->K,c->T,c->R, r0,d0,tau); if(col<(int)c->solution.size()) sum+=c->solution[col]; }
      val[0].val=sum; return 1;
    }

    case F_FLOW_IMPORT: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      if(!c->haveTrade){ val[0].val=0.0; return 1; }
      int r0=r-1, tau=(ts-1)*c->H+(hr-1);
      double sum=0.0;
      for(int o0=0; o0<c->R; ++o0){ if(o0==r0) continue; int base_o=o0*c->colsPerReg; int col=base_o+col_flow_od_tau(c->E,c->K,c->T,c->R, o0,r0,tau); if(col<(int)c->solution.size()) sum+=c->solution[col]; }
      val[0].val=sum; return 1;
    }

    case F_FLOW_OD: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int o=(int)(val[0].val+0.5), d=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||o<1||o>c->R||d<1||d>c->R||o==d||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      if(!c->haveTrade){ val[0].val=0.0; return 1; }
      int o0=o-1, d0=d-1, tau=(ts-1)*c->H+(hr-1), base_o=o0*c->colsPerReg;
      int col=base_o+col_flow_od_tau(c->E,c->K,c->T,c->R, o0,d0,tau);
      val[0].val = (col<(int)c->solution.size())? c->solution[col] : 0.0; return 1;
    }

    // Prices & Revenues
    case F_LP_PRICE: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int tau=(ts-1)*c->H+(hr-1); val[0].val=c->lmp[(r-1)*c->T + tau]; return 1;
    }

    case F_LP_CAPSLICE_PRICE: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int tau=(ts-1)*c->H+(hr-1); val[0].val=c->capdual[(r-1)*c->T + tau]; return 1;
    }

    case F_LP_REV_TECH: { // Energy revenue $/yr
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      val[0].val=c->rev_tech[(r-1)*c->E + (e-1)]; return 1;
    }

    // Energy revenue per MW-year using (Existing + quantized ΔInstalled)
    case F_LP_REV_EN_PERMW_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg;
      double exist = c->exist_cap[(r-1)*c->E + (e-1)];
      double b = c->solution[base+col_build_e(e-1)];
      double q = c->solution[base+col_retire_e(c->E,e-1)];
      double mppc = (e-1)<(int)c->mppc_E.size()? c->mppc_E[e-1] : 0.0;
      b = quantize_mppc(b, mppc, c->mppc_switch);
      q = quantize_mppc(q, mppc, c->mppc_switch);
      double installed = std::max(0.0, exist + b - q);
      if (installed <= 1e-12){ val[0].val = 0.0; return 1; }
      double revE = c->rev_tech[(r-1)*c->E + (e-1)]; // $/yr
      val[0].val = revE / installed; // $/MW-yr
      return 1;
    }

    // Capacity revenue per MW-year == capacity price tech
    case F_LP_REV_CAP_PERMW_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E || c->hours_SH.empty() || c->cf_RESH.empty()){ val[0].val=1e308; return 1; }
      double sum=0.0;
      for(int s=0;s<c->S;++s) for(int h=0; h<c->H; ++h){
        int tau = s*c->H + h;
        double dual = c->capdual[(r-1)*c->T + tau]; if (dual<=0.0 || !std::isfinite(dual)) continue;
        double hrs = c->hours_SH[tau];
        double cf  = c->cf_RESH[ ((r-1)*c->E + (e-1))*c->S*c->H + tau ];
        sum += dual * hrs * cf;
      }
      val[0].val = std::isfinite(sum)? sum : 0.0; return 1;
    }

    case F_LP_CAP_PRICE_TECH: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E || c->hours_SH.empty() || c->cf_RESH.empty()){ val[0].val=1e308; return 1; }
      double sum=0.0;
      for(int s=0;s<c->S;++s) for(int h=0; h<c->H; ++h){
        int tau = s*c->H + h;
        double dual = c->capdual[(r-1)*c->T + tau]; if (dual<=0.0 || !std::isfinite(dual)) continue;
        double hrs = c->hours_SH[tau];
        double cf  = c->cf_RESH[ ((r-1)*c->E + (e-1))*c->S*c->H + tau ];
        sum += dual * hrs * cf;
      }
      val[0].val = std::isfinite(sum)? sum : 0.0; return 1;
    }

    case F_LP_CAP_PRICE_REGION: {
      if(nval<2){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R || c->hours_SH.empty()){ val[0].val=1e308; return 1; }
      double sum=0.0;
      for(int s=0;s<c->S;++s) for(int h=0; h<c->H; ++h){
        int tau = s*c->H + h;
        double dual = c->capdual[(r-1)*c->T + tau]; if (dual<=0.0 || !std::isfinite(dual)) continue;
        sum += dual * c->hours_SH[tau];
      }
      val[0].val = std::isfinite(sum)? sum : 0.0; return 1;
    }

    // Storage getters
    case F_STOR_BUILD_CAP: {
      if(nval<2){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg;
      val[0].val = c->solution[base+col_stor_build(c->E,c->K,c->T,c->R,c->haveTrade)];
      return 1;
    }
    case F_STOR_CHARGE: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int tau=(ts-1)*c->H+(hr-1), base=(r-1)*c->colsPerReg;
      val[0].val = c->solution[base+col_stor_charge_tau(c->E,c->K,c->T,c->R,c->haveTrade,tau)];
      return 1;
    }
    case F_STOR_DISCHARGE: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int tau=(ts-1)*c->H+(hr-1), base=(r-1)*c->colsPerReg;
      val[0].val = c->solution[base+col_stor_discharge_tau(c->E,c->K,c->T,c->R,c->haveTrade,tau)];
      return 1;
    }
    case F_STOR_SOC: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int tau=(ts-1)*c->H+(hr-1), base=(r-1)*c->colsPerReg;
      val[0].val = c->solution[base+col_stor_soc_tau(c->E,c->K,c->T,c->R,c->haveTrade,tau)];
      return 1;
    }

    default: return 0;
  }
}
