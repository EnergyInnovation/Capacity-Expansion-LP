// lp_vensim_addon.cpp — HiGHS C API (multi-scenario, post-installed fix + quantize/repair)
//
// Highlights:
//  • Region-first layout; Build, Retire, Dispatch
//  • Reserve margin enforced on peak slices (imports count); reserve/energy prices exposed
//  • Region-specific costs: CAPEX[R,E], FOM_new[R,E], Fuel[R,E], VOM[R,E]
//  • CES / RPS with optional ACP slacks
//  • Per-(R,E) MaxBuild / CapMax
//  • Inter-regional trading with directional caps TransRR[R,R] (<=0 disables direction)
//  • Multi-scenario snapshots (scenario_id) and getters
//  • LMPs (demand duals), reserve duals (peak-only, clamped/rectified)
//  • Revenues and flows
//  • Capacity price getters ($/MW-yr = Σ dual × Hours × [× CF])
//  • Retirement trigger is boolean BelowFOMFlag[R,E] (0/1)
//  • FIX: LP_REV_EN_PERMW_INSTALLED divides by Installed_post = Existing + Build − Retire
//  • NEW: In-call Build/Retire **quantize + repair** using MinBlock_E[E] (0=disabled) & EnableMinBlock (0/1)
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
// If CRT mismatch, use /MD and the x64-windows triplet.

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

// ===== Current-solve globals =====
static std::mutex     g_mutex;
static vector<double> g_solution, g_cost_cap, g_cost_energy;
static int            g_status=-1, g_lastCode=-999;
static double         g_obj=0.0;
static bool           g_solved=false;
static int            g_R=0,g_S=0,g_H=0,g_E=0,g_T=0;
static int            g_colsPerReg=0,g_N_region=0,g_N_total=0;
static int            g_colSlackCES=-1,g_colSlackRPS=-1;
static bool           g_haveTrade=false;

// Reporting (last solve)
static vector<double> g_lmp;            // [R*T] demand duals ($/MWh)
static vector<double> g_capdual;        // [R*T] reserve duals ($/MWh) peak-only, clamped >=0
static vector<double> g_rev_tech;       // [R*E] energy rev Σ LMP*Gen ($/yr)
static vector<double> g_exist_cap;      // [R*E] snapshot of existing (MW)
static vector<double> g_installed_post; // [R*E] Installed_post = Existing + Build − Retire (MW)
// Inputs cached for capacity price calc
static vector<double> g_hours_SH;       // [S*H]
static vector<double> g_cf_RESH;        // [R*E*S*H]

// ===== Region block layout =====
// Per region: [Build E][Retire E][Gen E*T][Flows (R-1)*T?]
static inline int cols_per_region(int R,int E,int T,bool haveTrade){ return (2*E) + (E*T) + (haveTrade ? (R-1)*T : 0); }

static inline void reset_cache(int R,int E,int T,bool haveCES,bool haveRPS,bool haveTrade){
  g_haveTrade   = haveTrade;
  g_colsPerReg  = cols_per_region(R,E,T,haveTrade);
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
  /* keep g_hours_SH and g_cf_RESH as provided by F_SOLVE */
  g_exist_cap.clear(); g_installed_post.clear();
}

// ===== Multi-scenario snapshot =====
struct Context{
  int R=0,S=0,H=0,E=0,T=0;
  int colsPerReg=0,N_region=0,N_total=0;
  int colSlackCES=-1,colSlackRPS=-1;
  bool haveTrade=false;

  vector<double> solution, cost_cap, cost_energy;
  int status=-1,lastCode=-999;
  double obj=1e308;
  bool solved=false;

  vector<double> lmp, capdual, rev_tech, exist_cap, installed_post;
  vector<double> hours_SH, cf_RESH;
};
static std::unordered_map<int,Context> g_ctx;
static inline void save_context(int sid){
  Context c;
  c.R=g_R; c.S=g_S; c.H=g_H; c.E=g_E; c.T=g_T;
  c.colsPerReg=g_colsPerReg; c.N_region=g_N_region; c.N_total=g_N_total;
  c.colSlackCES=g_colSlackCES; c.colSlackRPS=g_colSlackRPS; c.haveTrade=g_haveTrade;
  c.solution=g_solution; c.cost_cap=g_cost_cap; c.cost_energy=g_cost_energy;
  c.status=g_status; c.lastCode=g_lastCode; c.obj=g_obj; c.solved=g_solved;
  c.lmp=g_lmp; c.capdual=g_capdual; c.rev_tech=g_rev_tech;
  c.exist_cap=g_exist_cap; c.installed_post=g_installed_post;
  c.hours_SH=g_hours_SH; c.cf_RESH=g_cf_RESH;
  g_ctx[sid]=std::move(c);
}
static inline const Context* get_ctx(int sid){ auto it=g_ctx.find(sid); return it==g_ctx.end()? nullptr : &it->second; }

// ===== Index helpers =====
static inline int idx_dem  (int r,int s,int h,int S,int H){ return ((r*S)+s)*H + h; }              // [R,S,H]
static inline int idx_hours(int s,int h,int H){ return s*H + h; }                                    // [S,H]
static inline int idx_cf   (int r,int e,int s,int h,int E,int S,int H){ return (((r*E+e)*S+s)*H+h);} // [R,E,S,H]
static inline int idx_mask (int s,int h,int H){ return s*H + h; }                                    // [S,H]
static inline int idx_re   (int r,int e,int E){ return r*E + e; }                                    // [R,E]

// Columns within region block
static inline int col_build_e (int e){ return e; }                         // 0..E-1
static inline int col_retire_e(int E,int e){ return E + e; }               // E..2E-1
static inline int col_gen_e_tau(int E,int e,int tau,int T){ return 2*E + e*T + tau; }
static inline int offset_trade(int E,int T){ return 2*E + E*T; }
static inline int dest_idx_for_origin(int r,int d){ return (d<r)? d : (d-1); }
static inline int col_flow_od_tau(int E,int T,int R,int r,int d,int tau){ return offset_trade(E,T) + dest_idx_for_origin(r,d)*T + tau; }

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

// ===== Core solver (with optional quantize+repair) =====
struct PassResult {
  bool ok=false;
  double obj=1e308;
  vector<double> col_val, row_dual;
};
static double solve_global_policies(
  // Demand & physics
  const double* Demand,             // [R,S,H]
  const double* CF,                 // [R,E,S,H]
  const double* Hours,              // [S,H]
  const double* PeakMask,           // [S,H] (1 at peaks else 0)

  // Region-specific costs (new builds / variable)
  const double* CAPEX_RE,           // [R,E]
  const double* FOM_new_RE,         // [R,E]
  const double* Fuel_RE,            // [R,E]
  const double* VOM_RE,             // [R,E]

  // Existing & caps
  const double* ExistingCap,        // [R,E]
  const double* MaxBuild,           // [R,E] (<=0 ignore)
  const double* CapMax,             // [R,E] (<=0 ignore)

  // Trading caps (<=0 disables that direction)
  const double* TransRR,            // [R,R]

  // CES / RPS
  const double* CES_q, double CES_rhs, double CES_ACP, // [E], scalars
  const double* RPS_q, double RPS_rhs, double RPS_ACP, // [E], scalars

  // Retirement economics
  const double* FOM_exist_RE,       // [R,E] $/MW-yr existing
  const double* RetireCost_RE,      // [R,E] $/MW-yr (annualized)
  const double* RevExist_RE,        // [R,E] $/MW-yr observed
  const double* RetireResponse_RE,  // [R,E] MW per ($/MW-yr)
  const double* BelowFOMFlag_RE,    // [R,E] 0/1
  double FOM_cover_mult,            // 0..1

  // Reserve
  double ReserveMargin,

  // Quantize controls
  const double* MinBlock_E,         // [E] MW (0 => disabled)
  int EnableMinBlock,               // 0/1

  int R,int S,int H,int E
){
  const int T=S*H;

  auto is_peak = [&](int s,int h0)->bool{
    if(!PeakMask) return true;
    return PeakMask[idx_mask(s,h0,H)] > 0.5;
  };

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

  const bool useCES=(CES_rhs>0.0)||(CES_ACP>0.0);
  const bool addCES=(CES_ACP>0.0);
  const bool useRPS=(RPS_rhs>0.0)||(RPS_ACP>0.0);
  const bool addRPS=(RPS_ACP>0.0);

  reset_cache(R,E,T, addCES, addRPS, haveTrade);
  g_S=S; g_H=H;

  // ---------- a lambda to run one LP pass (optionally fixing Build/Retire) ----------
  auto run_one_pass = [&](const vector<double>* fixBuild, const vector<double>* fixRet)->PassResult {
    // ----- Rows -----
    const int ROW_DEM0=0;                       // R*T
    const int ROW_CAP0=ROW_DEM0 + R*T;          // R*E*T
    int rowCount = ROW_CAP0 + R*E*T;
    const bool useRes=(ReserveMargin>0.0);
    const int ROW_RSV0=rowCount; if(useRes) rowCount+=R*T;
    const int ROW_CES=rowCount;  if(useCES) rowCount+=1;
    const int ROW_RPS=rowCount;  if(useRPS) rowCount+=1;

    // ----- Columns -----
    const double INF=1e30;
    vector<double> col_cost(g_N_total,0.0), col_lo(g_N_total,0.0), col_hi(g_N_total,INF);

    for(int r=0;r<R;++r){
      const int base=r*g_colsPerReg;

      // Build
      for(int e0=0;e0<E;++e0){
        int cB=base+col_build_e(e0);
        double capex=CAPEX_RE[idx_re(r,e0,E)];
        double fomN =FOM_new_RE[idx_re(r,e0,E)];
        col_cost[cB]=capex + fomN;
        col_lo[cB]=0.0;

        double ubB=INF;
        if(MaxBuild){ double mb=MaxBuild[idx_re(r,e0,E)]; if(mb>0.0) ubB=std::min(ubB,mb); }
        if(CapMax){ double cm=CapMax[idx_re(r,e0,E)]; if(cm>0.0){ double ex=ExistingCap[idx_re(r,e0,E)]; ubB=std::min(ubB,std::max(0.0, cm - ex)); } }

        if (fixBuild){
          double fb = (*fixBuild)[idx_re(r,e0,E)];
          col_lo[cB]=fb; col_hi[cB]=fb;
        }else{
          col_hi[cB]=ubB;
        }
      }

      // Retire (≤ existing, economics UB)
      for(int e0=0;e0<E;++e0){
        int cR=base+col_retire_e(E,e0);
        col_lo[cR]=0.0;
        if (fixRet){
          double fr = (*fixRet)[idx_re(r,e0,E)];
          col_lo[cR]=fr; col_hi[cR]=fr;
        }else{
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
        }
        double retireCost=RetireCost_RE[idx_re(r,e0,E)];
        double Fexist=FOM_exist_RE[idx_re(r,e0,E)];
        col_cost[cR]=retireCost - Fexist;
      }

      // Gen
      for(int e0=0;e0<E;++e0){
        double vc = Fuel_RE[idx_re(r,e0,E)] + VOM_RE[idx_re(r,e0,E)];
        for(int tau=0;tau<T;++tau){
          int cG=base+col_gen_e_tau(E,e0,tau,T);
          col_cost[cG]=vc; col_lo[cG]=0.0; col_hi[cG]=INF;
        }
      }

      // Flows
      if(haveTrade){
        for(int d=0; d<R; ++d){
          if(d==r) continue;
          bool ok = allow[r*R + d];
          for(int tau=0;tau<T;++tau){
            int cF=base+col_flow_od_tau(E,T,R, r,d,tau);
            col_cost[cF]=0.0; col_lo[cF]=0.0; col_hi[cF]= ok? TransRR[r*R + d] : 0.0;
          }
        }
      }
    }

    // Slacks
    if(g_colSlackCES>=0){ col_cost[g_colSlackCES]=CES_ACP; col_lo[g_colSlackCES]=0.0; col_hi[g_colSlackCES]=INF; }
    if(g_colSlackRPS>=0){ col_cost[g_colSlackRPS]=RPS_ACP; col_lo[g_colSlackRPS]=0.0; col_hi[g_colSlackRPS]=INF; }

    // ----- Row bounds -----
    vector<double> row_lo(rowCount,-INF), row_hi(rowCount,INF);

    // Demand (=)
    for(int r=0;r<R;++r)
      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0, row=ROW_DEM0 + r*T + tau;
          double rhs=Demand[idx_dem(r,s,h0,S,H)];
          row_lo[row]=rhs; row_hi[row]=rhs;
        }

    // Capacity (>=): -Gen + CF*hrs*(Build - Retire) >= -CF*hrs*Existing
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

    // Reserve rows (>=) only at peak slices. RHS=(1+RM)*Demand - Σ CF*hrs*Existing
    const bool useReserve=(ReserveMargin>0.0);
    if(useReserve){
      for(int r=0;r<R;++r)
        for(int s=0;s<S;++s)
          for(int h0=0;h0<H;++h0){
            int tau=s*H+h0, row=ROW_RSV0 + r*T + tau;
            if(!(PeakMask && PeakMask[idx_mask(s,h0,H)] > 0.5)){ row_lo[row]=-INF; row_hi[row]=INF; continue; }
            double hrs=Hours[idx_hours(s,h0,H)];
            double rhs=(1.0+ReserveMargin)*Demand[idx_dem(r,s,h0,S,H)];
            for(int e0=0;e0<E;++e0) rhs -= CF[idx_cf(r,e0,s,h0,E,S,H)] * hrs * ExistingCap[idx_re(r,e0,E)];
            row_lo[row]=rhs; row_hi[row]=INF;
          }
    }

    // CES / RPS (>=)
    if(useCES){ row_lo[ROW_CES]=std::max(0.0,CES_rhs); row_hi[ROW_CES]=INF; }
    if(useRPS){ row_lo[ROW_RPS]=std::max(0.0,RPS_rhs); row_hi[ROW_RPS]=INF; }

    // ----- Triplets -----
    vector<Triplet> Tpls;
    Tpls.reserve((size_t)R*E*T*4 + (useReserve? (size_t)R*T*E : 0)
                 + (useCES? (size_t)R*E*T + (g_colSlackCES>=0?1:0) : 0)
                 + (useRPS? (size_t)R*E*T + (g_colSlackRPS>=0?1:0) : 0)
                 + (size_t)R*T*E + (g_haveTrade? (size_t)R*(R-1)*T*2 : 0));

    for(int r=0;r<R;++r){
      int base=r*g_colsPerReg;

      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0; double hrs=Hours[idx_hours(s,h0,H)];

          // Demand: Σe Gen + inflows − outflows = Demand
          int rowD=ROW_DEM0 + r*T + tau;
          for(int e0=0;e0<E;++e0) Tpls.push_back({rowD, base+col_gen_e_tau(E,e0,tau,T), 1.0});

          if(g_haveTrade){
            // inflows from any origin o
            for(int o=0;o<R;++o){
              if(o==r) continue;
              if(!(TransRR && TransRR[o*R + r] > 0.0)) continue;
              int base_o=o*g_colsPerReg;
              Tpls.push_back({rowD, base_o+col_flow_od_tau(E,T,R, o,r,tau), 1.0});
            }
            // outflows to any d
            for(int d=0; d<R; ++d){
              if(d==r) continue;
              if(!(TransRR && TransRR[r*R + d] > 0.0)) continue;
              Tpls.push_back({rowD, base+col_flow_od_tau(E,T,R, r,d,tau), -1.0});
            }
          }

          // Capacity rows: -Gen + CF*hrs*(Build - Retire) >= -CF*hrs*Existing
          for(int e0=0;e0<E;++e0){
            double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
            int rowC=ROW_CAP0 + ((r*E+e0)*T + tau);
            Tpls.push_back({rowC, base+col_gen_e_tau(E,e0,tau,T), -1.0});
            if(cf!=0.0 && hrs!=0.0){
              Tpls.push_back({rowC, base+col_build_e(e0),  cf*hrs});
              Tpls.push_back({rowC, base+col_retire_e(E,e0), -cf*hrs});
            }
          }

          // Reserve rows (peak): local build-retire contributes; imports contribute +1
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
                Tpls.push_back({rowR, base_o+col_flow_od_tau(E,T,R, o,r,tau), 1.0});
              }
            }
          }

          // CES / RPS rows
          if(useCES){ int rowCES=ROW_CES; for(int e0=0;e0<E;++e0){ double q=CES_q[e0]; if(q!=0.0) Tpls.push_back({rowCES, base+col_gen_e_tau(E,e0,tau,T), q}); } }
          if(useRPS){ int rowRPS=ROW_RPS; for(int e0=0;e0<E;++e0){ double q=RPS_q[e0]; if(q!=0.0) Tpls.push_back({rowRPS, base+col_gen_e_tau(E,e0,tau,T), q}); } }
        }
    }

    if(useCES && g_colSlackCES>=0) Tpls.push_back({ROW_CES,g_colSlackCES,1.0});
    if(useRPS && g_colSlackRPS>=0) Tpls.push_back({ROW_RPS,g_colSlackRPS,1.0});

    // Convert & solve
    vector<int> astart_i,aindex_i; vector<double> avalue;
    triplets_to_csc(rowCount,g_N_total,Tpls,astart_i,aindex_i,avalue);
    vector<HighsInt> astart(astart_i.begin(),astart_i.end()), aindex(aindex_i.begin(),aindex_i.end());
    HighsInt nnz=(HighsInt)avalue.size();
    void* h=Highs_create();
    const HighsInt a_format=1, sense=1; const double obj_offset=0.0;

    int rc = Highs_passLp(h,(HighsInt)g_N_total,(HighsInt)rowCount,nnz, a_format,sense,obj_offset,
                          g_N_total? &col_cost[0]:nullptr, g_N_total? &col_lo[0]:nullptr, g_N_total? &col_hi[0]:nullptr,
                          rowCount? &row_lo[0]:nullptr, rowCount? &row_hi[0]:nullptr,
                          astart.data(), aindex.data(), avalue.data());
    if(rc==0) rc=Highs_run(h);

    PassResult pr;
    pr.ok = (rc==0);
    if(pr.ok){
      pr.obj=Highs_getObjectiveValue(h);
      pr.col_val.assign(g_N_total,0.0);
      vector<double> col_dual(g_N_total,0.0);
      vector<double> row_val(std::max(1,rowCount),0.0);
      pr.row_dual.assign(std::max(1,rowCount),0.0);
      if(Highs_getSolution(h, pr.col_val.data(), col_dual.data(), row_val.data(), pr.row_dual.data())!=0){
        pr.ok=false;
      }
    }
    Highs_destroy(h);
    return pr;
  };

  // ---- Pass 1: free Build/Retire ----
  PassResult pass1 = run_one_pass(nullptr, nullptr);
  g_lastCode = pass1.ok ? 0 : 1;
  g_status = pass1.ok ? 0 : -1;   // placeholder; later we set 7 for optimal
  g_obj      = 1e308; g_solved=false;
  if(!pass1.ok) return 1e308;

  // Optionally quantize + repair
  bool do_repair=false;
  vector<double> qB, qR; // [R*E]
  if (EnableMinBlock && MinBlock_E){
    const double EPS=1e-9;
    qB.assign((size_t)R*E, 0.0);
    qR.assign((size_t)R*E, 0.0);
    for(int r=0;r<R;++r){
      int base=r*g_colsPerReg;
      for(int e0=0;e0<E;++e0){
        double m = MinBlock_E[e0];
        if(!(m>0.0) || !std::isfinite(m)) m = 0.0; // 0 => disabled
        double b = pass1.col_val[base + col_build_e(e0)];
        double rMW= pass1.col_val[base + col_retire_e(E,e0)];
        if (m>0.0){
          double qb = m * std::floor((b + EPS)/m);
          double qr = m * std::floor((rMW + EPS)/m);
          if (std::fabs(qb - b) > 1e-8 || std::fabs(qr - rMW) > 1e-8) do_repair = true;
          qB[idx_re(r,e0,E)] = qb;
          qR[idx_re(r,e0,E)] = qr;
        }else{
          qB[idx_re(r,e0,E)] = b;
          qR[idx_re(r,e0,E)] = rMW;
        }
      }
    }
  }

  PassResult passF = pass1;
  if (do_repair){
    passF = run_one_pass(&qB, &qR);
    if(!passF.ok){
      // Fall back to the unconstrained pass if repair unexpectedly fails
      passF = pass1;
    }
  }

  // ----- store final solution & reporting -----
  g_solution = std::move(passF.col_val);

  // LMPs from demand duals
  g_lmp.assign(R*T,0.0);
  {
    // To reconstruct row indices we need the same structure as in run_one_pass
    const int ROW_DEM0=0;
    const int ROW_CAP0=ROW_DEM0 + R*T;
    const bool useRes=(ReserveMargin>0.0);
    const int ROW_RSV0=ROW_CAP0 + R*E*T;
    // row_dual from the pass used:
    const vector<double>& row_dual = passF.row_dual;
    for(int r=0;r<R;++r) for(int tau=0;tau<T;++tau){
      int row=ROW_DEM0 + r*T + tau;
      g_lmp[r*T + tau] = (row < (int)row_dual.size()) ? row_dual[row] : 0.0;
    }

    // Reserve duals — peak-only, clamp, rectify
    g_capdual.assign(R*T, 0.0);
    auto clamp_dual = [](double d)->double {
      if (!std::isfinite(d)) return 0.0;
      if (d < 0.0) d = -d;
      const double DMAX = 1e6; if (d > DMAX) d = DMAX;
      return d;
    };
    if (useRes){
      for(int r=0;r<R;++r) for(int s=0;s<S;++s) for(int h0=0;h0<H;++h0){
        int tau=s*H + h0;
        int row=ROW_RSV0 + r*T + tau;
        double d = (is_peak(s,h0) && row<(int)row_dual.size()) ? clamp_dual(row_dual[row]) : 0.0;
        g_capdual[r*T + tau] = d;
      }
    }
  }

  // Energy revenue per (r,e)
  g_rev_tech.assign(R*E,0.0);
  for(int r=0;r<R;++r){
    int base=r*g_colsPerReg;
    for(int e0=0;e0<E;++e0){
      double rev=0.0;
      for(int tau=0;tau<T;++tau){
        int cG=base+col_gen_e_tau(E,e0,tau,T);
        rev += g_lmp[r*T + tau] * g_solution[cG];
      }
      g_rev_tech[r*E + e0]=rev;
    }
  }

  // Post-decision installed capacity (for per-MW denominators)
  g_installed_post.assign(R*E, 0.0);
  for(int r=0;r<R;++r){
    int base=r*g_colsPerReg;
    for(int e0=0;e0<E;++e0){
      const double exist  = ExistingCap[idx_re(r,e0,E)];
      const double build  = g_solution[base + col_build_e(e0)];
      const double retire = g_solution[base + col_retire_e(E,e0)];
      g_installed_post[r*E + e0] = std::max(0.0, exist + build - retire);
    }
  }

  // Keep a clean snapshot for the saved context
  g_exist_cap.assign(R*E, 0.0);
  for(int r=0;r<R;++r)
    for(int e0=0;e0<E;++e0)
      g_exist_cap[r*E + e0] = ExistingCap[idx_re(r,e0,E)];

  // Regional cost breakout (capacity bucket includes retire effects)
  for(int r=0;r<R;++r){
    double cc=0.0, ce=0.0; int base=r*g_colsPerReg;
    for(int e0=0;e0<E;++e0){
      int cB=base+col_build_e(e0), cR=base+col_retire_e(E,e0);
      double capex = CAPEX_RE    [idx_re(r,e0,E)];
      double fomN  = FOM_new_RE  [idx_re(r,e0,E)];
      double retire= RetireCost_RE[idx_re(r,e0,E)];
      double Fexist= FOM_exist_RE[idx_re(r,e0,E)];
      cc += (capex + fomN) * g_solution[cB] + (retire - Fexist) * g_solution[cR];
      double vc=Fuel_RE[idx_re(r,e0,E)] + VOM_RE[idx_re(r,e0,E)];
      for(int tau=0;tau<T;++tau) ce += vc * g_solution[base+col_gen_e_tau(E,e0,tau,T)];
    }
    g_cost_cap[r]=cc; g_cost_energy[r]=ce;
  }

  g_obj=passF.obj; g_solved=std::isfinite(g_obj); g_status = 7; // 7=optimal in HiGHS (for reference)
  return g_solved? g_obj : 1e308;
}

// ===== Vensim glue =====
static const int EXTERN_VCODE=62051;             // DO NOT CHANGE
extern "C" __declspec(dllexport) int VEFCC version_info(){ return EXTERN_VCODE; }

enum {
  F_SOLVE=1101,
  F_CAP_ADD, F_RETIRE, F_CAP_INST, F_GEN,
  F_COST_CAP, F_COST_ENERGY, F_LP_OBJ, F_LP_STATUS, F_LP_CODE,
  F_FLOW_EXPORT, F_FLOW_IMPORT, F_FLOW_OD,
  // Reporting
  F_LP_PRICE, F_LP_CAPSLICE_PRICE, F_LP_REV_TECH,
  F_LP_REV_EN_PERMW_INST,   // energy revenue per MW-yr (post-decision installed)
  F_LP_REV_CAP_PERMW_INST,  // capacity revenue per MW-yr (= cap price tech)
  F_LP_CAP_PRICE_TECH, F_LP_CAP_PRICE_REGION
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
      // Vectors (21): Demand, CF, Hours, PeakMask, CAPEX, FOM_new, Fuel, VOM,
      //               ExistingCap, MaxBuild, CapMax, TransRR,
      //               CES_q, RPS_q, FOM_exist, RetireCost, RevExist, RetireResponse, BelowFOMFlag,
      //               MinBlock_E
      // Scalars: CES_rhs, CES_ACP, RPS_rhs, RPS_ACP, FOM_cover_mult, ReserveMargin, R,S,H,E, scenario_id, EnableMinBlock
      *arglist=(char*)"{Demand},{Capacity_factor},{Hours},{PeakMask},{CAPEX},{FOM_new},{Fuel},{VOM},{ExistingCap},{MaxBuild},{CapMax},{TransRR},{CES_qualifying},{RPS_qualifying},{FOM_exist},{RetireCost},{RevExist},{RetireResponse},{BelowFOMFlag},{MinBlock_E},CES_rhs,CES_ACP,RPS_rhs,RPS_ACP,FOM_cover_mult,ReserveMargin,R,S,H,E,scenario_id,EnableMinBlock";
      *num_args=32; *num_vector=21; *func_index=F_SOLVE; return 1;

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

    default: return 0;
  }
}

extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch(funcid){

    case F_SOLVE: {
      if(nval<32){ val[0].val=1e308; return 0; }

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

      const double* MinBlockE = val[19].vec->firstval; // [E]

      double CES_rhs = val[20].val, CES_ACP = val[21].val;
      double RPS_rhs = val[22].val, RPS_ACP = val[23].val;
      double FOM_cover_mult = val[24].val;
      double ReserveMargin  = val[25].val;

      int R=(int)(val[26].val+0.5), S=(int)(val[27].val+0.5), H=(int)(val[28].val+0.5), E=(int)(val[29].val+0.5);
      int scenario_id        = (int)(val[30].val+0.5);
      int EnableMinBlock     = (int)(val[31].val+0.5);

      // Persist inputs needed for post-processing
      g_exist_cap.assign(R*E,0.0);
      for(int r=0;r<R;++r) for(int e=0;e<E;++e) g_exist_cap[r*E+e]=XCap[r*E+e];

      g_hours_SH.assign(S*H,0.0);
      for(int s=0;s<S;++s) for(int h=0;h<H;++h) g_hours_SH[s*H+h]=Hours[s*H+h];

      g_cf_RESH.assign((size_t)R*E*S*H,0.0);
      for(int r=0;r<R;++r) for(int e=0;e<E;++e) for(int s=0;s<S;++s) for(int h=0;h<H;++h)
        g_cf_RESH[ ((r*E + e)*S + s)*H + h ] = CF[ ((r*E + e)*S + s)*H + h ];

      val[0].val = solve_global_policies(
        Dem,CF,Hours,PeakM,
        CAPEX,FOMnew,Fuel,VOM,
        XCap,MaxB,CapMx,
        Trans,
        CES_q,CES_rhs,CES_ACP, RPS_q,RPS_rhs,RPS_ACP,
        FOMexist,RetCost,RevExist,RetResp,FlagBF,
        FOM_cover_mult,
        ReserveMargin,
        MinBlockE, EnableMinBlock,
        R,S,H,E
      );
      save_context(scenario_id);
      return 1;
    }

    // --- Basic getters ---
    case F_CAP_ADD: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; val[0].val=c->solution[base+col_build_e(e-1)]; return 1;
    }

    case F_RETIRE: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; val[0].val=c->solution[base+col_retire_e(c->E,e-1)]; return 1;
    }

    // ΔInstalled (Build − Retire)
    case F_CAP_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg;
      double build=c->solution[base+col_build_e(e-1)];
      double retire=c->solution[base+col_retire_e(c->E,e-1)];
      val[0].val = build - retire;
      return 1;
    }

    case F_GEN: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; int tau=(ts-1)*c->H+(hr-1);
      val[0].val=c->solution[base+col_gen_e_tau(c->E,e-1,tau,c->T)]; return 1;
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
      for(int d0=0; d0<c->R; ++d0){ if(d0==r0) continue; int col=base_r+col_flow_od_tau(c->E,c->T,c->R, r0,d0,tau); if(col<(int)c->solution.size()) sum+=c->solution[col]; }
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
      for(int o0=0; o0<c->R; ++o0){ if(o0==r0) continue; int base_o=o0*c->colsPerReg; int col=base_o+col_flow_od_tau(c->E,c->T,c->R, o0,r0,tau); if(col<(int)c->solution.size()) sum+=c->solution[col]; }
      val[0].val=sum; return 1;
    }

    case F_FLOW_OD: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int o=(int)(val[0].val+0.5), d=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||o<1||o>c->R||d<1||d>c->R||o==d||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      if(!c->haveTrade){ val[0].val=0.0; return 1; }
      int o0=o-1, d0=d-1, tau=(ts-1)*c->H+(hr-1), base_o=o0*c->colsPerReg;
      int col=base_o+col_flow_od_tau(c->E,c->T,c->R, o0,d0,tau);
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

    // Energy revenue per MW-year using Installed_post (robust vs. “no build” years)
    case F_LP_REV_EN_PERMW_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E || c->installed_post.empty()){ val[0].val=1e308; return 1; }
      double installed = c->installed_post[(r-1)*c->E + (e-1)];
      if (installed <= 1e-12){ val[0].val = 0.0; return 1; }
      double revE = c->rev_tech[(r-1)*c->E + (e-1)]; // $/yr
      val[0].val = revE / installed; // $/MW-yr
      return 1;
    }

    // Capacity revenue per MW-year == capacity price for that tech
    case F_LP_REV_CAP_PERMW_INST: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E || c->hours_SH.empty() || c->cf_RESH.empty()){ val[0].val=1e308; return 1; }
      double sum=0.0;
      for(int s=0;s<c->S;++s) for(int h=0; h<c->H; ++h){
        int tau = s*c->H + h;
        double dual = c->capdual[(r-1)*c->T + tau];
        if (dual<=0.0 || !std::isfinite(dual)) continue;
        double hrs = c->hours_SH[tau];
        double cf  = c->cf_RESH[ ((r-1)*c->E + (e-1))*c->S*c->H + tau ];
        sum += dual * hrs * cf;
      }
      if (!std::isfinite(sum)) sum = 0.0;
      val[0].val = sum; // $/MW-yr
      return 1;
    }

    // Capacity price ($/MW-yr) and region-average ($/MW-yr)
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

    default: return 0;
  }
}
