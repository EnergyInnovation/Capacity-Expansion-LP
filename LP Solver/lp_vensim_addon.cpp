// lp_vensim_addon.cpp — HiGHS C API (multi-scenario only API)
// Features:
//  - Region-first layout
//  - Capacity build + dispatch with reserve constraint
//  - CES / RPS rows with optional ACP (slacks priced at CES_ACP / RPS_ACP)
//  - Per-(region,tech) build caps (MaxBuild[R,E]) and installed caps (CapMax[R,E])
//  - Inter-regional trading via pairwise directional flows Flow[r->d,τ] with caps TransRR[R,R]
//  - Multi-scenario snapshots: ALL getters take scenario_id; no global (last-solve) getters
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
// If you hit CRT mismatch, swap /MT -> /MD to match your highs.lib.

#define NOMINMAX
#include <windows.h>
#include <vector>
#include <mutex>
#include <algorithm>
#include <cstdint>
#include <cmath>
#include <unordered_map>

#include <highs/interfaces/highs_c_api.h>

#if defined(_MSC_VER)
  #define VEFCC __stdcall
#else
  #define VEFCC
#endif

using std::vector;

// ===== Vensim ABI-ish types =====
typedef double COMPREAL;
typedef struct {
  COMPREAL*       vals;
  const COMPREAL* firstval;
  const void*     dim_info;
  const char*     varname;
} VECTOR_ARG;

typedef union {
  COMPREAL    val;
  VECTOR_ARG* vec;
  void*       tab;
  char*       literal;
  void*       constmat;
  void*       datamat;
} VV;

// ---------- Working buffers (for current solve only) ----------
static std::mutex          g_mutex;
static vector<double>      g_solution;    // size N_total
static vector<double>      g_cost_cap;    // size R
static vector<double>      g_cost_energy; // size R
static int                 g_status   = -1;
static int                 g_lastCode = -999;
static double              g_obj      = 0.0;
static bool                g_solved   = false;

static int g_R=0, g_S=0, g_H=0, g_E=0;
static int g_T=0, g_colsPerReg=0, g_N_region=0, g_N_total=0;
static int g_colSlackCES=-1, g_colSlackRPS=-1;

static inline void reset_cache(int R,int E,int T,bool haveCESslack,bool haveRPSslack,bool haveTrade){
  g_colsPerReg = E*(T+1) + (haveTrade ? (R-1)*T : 0);
  g_N_region   = R*g_colsPerReg;
  g_N_total    = g_N_region + (haveCESslack?1:0) + (haveRPSslack?1:0);
  g_solution.assign(g_N_total,0.0);
  g_cost_cap.assign(R,0.0);
  g_cost_energy.assign(R,0.0);
  g_status=-1; g_lastCode=-999; g_obj=0.0; g_solved=false;
  g_R=R; g_E=E; g_T=T;
  int off=g_N_region;
  g_colSlackCES = haveCESslack? off++ : -1;
  g_colSlackRPS = haveRPSslack? off++ : -1;
}

// ---------- Multi-scenario snapshots ----------
struct Context{
  int R=0,S=0,H=0,E=0,T=0, colsPerReg=0,N_region=0,N_total=0;
  int colSlackCES=-1,colSlackRPS=-1;
  vector<double> solution, cost_cap, cost_energy;
  int status=-1,lastCode=-999;
  double obj=1e308;
  bool solved=false;
};
static std::unordered_map<int,Context> g_ctx; // scenario_id -> snapshot

static inline void save_context(int sid){
  Context c;
  c.R=g_R; c.S=g_S; c.H=g_H; c.E=g_E; c.T=g_T;
  c.colsPerReg=g_colsPerReg; c.N_region=g_N_region; c.N_total=g_N_total;
  c.colSlackCES=g_colSlackCES; c.colSlackRPS=g_colSlackRPS;
  c.solution=g_solution; c.cost_cap=g_cost_cap; c.cost_energy=g_cost_energy;
  c.status=g_status; c.lastCode=g_lastCode; c.obj=g_obj; c.solved=g_solved;
  g_ctx[sid]=std::move(c);
}
static inline const Context* get_ctx(int sid){
  auto it=g_ctx.find(sid);
  return it==g_ctx.end()?nullptr:&it->second;
}

// ---------- Index helpers ----------
static inline int idx_dem(int r,int s,int h,int S,int H){ return ((r*S)+s)*H + h; }       // [R,S,H]
static inline int idx_hours(int s,int h,int H){ return s*H + h; }                          // [S,H]
static inline int idx_cf(int r,int e,int s,int h,int E,int S,int H){ return (((r*E+e)*S+s)*H+h);} // [R,E,S,H]
static inline int idx_xcap(int r,int e,int E){ return r*E + e; }                           // [R,E]

// Columns in region block
static inline int col_build_e(int e){ return e; }                                         // 0..E-1
static inline int col_gen_e_tau(int E,int e,int tau,int T){ return E + e*T + tau; }       // E..E+E*T-1

// Trading columns: origin r block contains flows to all d!=r
static inline int offset_trade(int E,int T){ return E*(T+1); }
static inline int dest_idx_for_origin(int r,int d){ return (d<r)?d:(d-1); }               // 0..R-2
static inline int col_flow_od_tau(int E,int T,int R,int r,int d,int tau){
  return offset_trade(E,T) + dest_idx_for_origin(r,d)*T + tau;
}

// ---------- Triplet -> CSC ----------
struct Triplet{ int row; int col; double val; };
static void triplets_to_csc(
  int nRows,int nCols,const vector<Triplet>& t,
  vector<int>& astart, vector<int>& aindex, vector<double>& avalue){
  astart.assign(nCols+1,0);
  for(const auto& x:t){ if(x.col>=0&&x.col<nCols) astart[x.col+1]++; }
  for(int c=0;c<nCols;++c) astart[c+1]+=astart[c];
  const int nnz=(int)t.size();
  aindex.assign(nnz,0); avalue.assign(nnz,0.0);
  vector<int> next=astart;
  for(const auto& x:t){ int p=next[x.col]++; aindex[p]=x.row; avalue[p]=x.val; }
}

// ---------- Core solve ----------
static double solve_global_policies(
  const double* Demand,
  const double* CF,
  const double* Hours,
  const double* CAPEX,
  const double* FOM,
  const double* Fuel,
  const double* VOM,
  const double* ExistingCap,
  const double* MaxBuild,
  const double* CapMax,
  const double* TransRR,
  const double* CES_q, double CES_rhs, double CES_ACP,
  const double* RPS_q, double RPS_rhs, double RPS_ACP,
  double ReserveMargin, int R,int S,int H,int E
){
  const int T=S*H;
  const bool useCES=(CES_rhs>0.0)||(CES_ACP>0.0);
  const bool addCES=(CES_ACP>0.0);
  const bool useRPS=(RPS_rhs>0.0)||(RPS_ACP>0.0);
  const bool addRPS=(RPS_ACP>0.0);
  const bool haveTrade=true;

  reset_cache(R,E,T,addCES,addRPS,haveTrade); g_S=S; g_H=H;

  // Rows
  const int ROW_DEM0=0;                 // R*T
  const int ROW_CAP0=ROW_DEM0+R*T;      // R*E*T
  int rowCount = ROW_CAP0 + R*E*T;
  const bool useRes=(ReserveMargin>0.0);
  const int ROW_RSV0=rowCount; if(useRes) rowCount+=R*T;
  const int ROW_CES=rowCount;  if(useCES) rowCount+=1;
  const int ROW_RPS=rowCount;  if(useRPS) rowCount+=1;

  // Columns
  const double INF=1e30; vector<double> col_cost(g_N_total,0.0), col_lo(g_N_total,0.0), col_hi(g_N_total,INF);

  // Build, Gen, Flow columns & bounds
  for(int r=0;r<R;++r){
    const int base=r*g_colsPerReg;
    // Build vars
    for(int e0=0;e0<E;++e0){
      const int c=base+col_build_e(e0);
      col_cost[c]=CAPEX[e0]+FOM[e0];
      col_lo[c]=0.0;
      double ub=INF;
      if(MaxBuild){
        double mb=MaxBuild[idx_xcap(r,e0,E)];
        if(mb>0.0) ub=std::min(ub,mb);
      }
      if(CapMax){
        double cm=CapMax[idx_xcap(r,e0,E)];
        if(cm>0.0){
          double ex=ExistingCap[idx_xcap(r,e0,E)];
          ub=std::min(ub,std::max(0.0,cm-ex));
        }
      }
      col_hi[c]=ub;
    }
    // Gen vars
    for(int e0=0;e0<E;++e0){
      const double vc=Fuel[e0]+VOM[e0];
      for(int tau=0;tau<T;++tau){
        int c=base+col_gen_e_tau(E,e0,tau,T);
        col_cost[c]=vc; col_lo[c]=0.0; col_hi[c]=INF;
      }
    }
    // Flow vars
    for(int d=0; d<R; ++d){
      if(d==r) continue;
      for(int tau=0;tau<T;++tau){
        int c=base+col_flow_od_tau(E,T,R,r,d,tau);
        col_cost[c]=0.0; col_lo[c]=0.0;
        double ub=INF;
        if(TransRR){
          double cap=TransRR[r*R+d];
          if(cap>0.0) ub=cap;
        }
        col_hi[c]=ub;
      }
    }
  }

  // Global slack columns
  if(g_colSlackCES>=0){ col_cost[g_colSlackCES]=CES_ACP; col_lo[g_colSlackCES]=0.0; col_hi[g_colSlackCES]=INF; }
  if(g_colSlackRPS>=0){ col_cost[g_colSlackRPS]=RPS_ACP; col_lo[g_colSlackRPS]=0.0; col_hi[g_colSlackRPS]=INF; }

  // Row bounds
  vector<double> row_lo(rowCount,-INF), row_hi(rowCount,INF);

  // Demand (=)
  for(int r=0;r<R;++r)
    for(int s=0;s<S;++s)
      for(int h0=0;h0<H;++h0){
        int tau=s*H+h0;
        int row=ROW_DEM0+r*T+tau;
        double rhs=Demand[idx_dem(r,s,h0,S,H)];
        row_lo[row]=rhs; row_hi[row]=rhs;
      }

  // Capacity (>=): -g + CF*hrs*Build >= -CF*hrs*Existing
  for(int r=0;r<R;++r)
    for(int e0=0;e0<E;++e0){
      double ex=ExistingCap[idx_xcap(r,e0,E)];
      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0; double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
          double hrs=Hours[idx_hours(s,h0,H)];
          int row=ROW_CAP0+((r*E+e0)*T+tau);
          row_lo[row]= -cf*hrs*ex; row_hi[row]=INF;
        }
    }

  // Reserve (>=): (1+RM)*Demand - sum(CF*hrs*Existing) <= sum(CF*hrs*Build)
  if(useRes){
    for(int r=0;r<R;++r)
      for(int s=0;s<S;++s)
        for(int h0=0;h0<H;++h0){
          int tau=s*H+h0;
          double rhs=(1.0+ReserveMargin)*Demand[idx_dem(r,s,h0,S,H)];
          double hrs=Hours[idx_hours(s,h0,H)];
          for(int e0=0;e0<E;++e0){
            rhs -= CF[idx_cf(r,e0,s,h0,E,S,H)]*hrs*ExistingCap[idx_xcap(r,e0,E)];
          }
          int row=ROW_RSV0+r*T+tau; row_lo[row]=rhs; row_hi[row]=INF;
        }
  }

  // CES / RPS (>=)
  if(useCES){ row_lo[ROW_CES]=std::max(0.0,CES_rhs); row_hi[ROW_CES]=INF; }
  if(useRPS){ row_lo[ROW_RPS]=std::max(0.0,RPS_rhs); row_hi[ROW_RPS]=INF; }

  // Matrix triplets
  vector<Triplet> Tpls;
  Tpls.reserve(
    (size_t)R*E*T*3
    + (useRes? (size_t)R*T*E : 0)
    + (useCES? (size_t)R*E*T + (g_colSlackCES>=0?1:0) : 0)
    + (useRPS? (size_t)R*E*T + (g_colSlackRPS>=0?1:0) : 0)
    + (size_t)R*T*E
    + (size_t)R*(R-1)*T*2
  );

  for(int r=0;r<R;++r){
    const int base=r*g_colsPerReg;
    for(int s=0;s<S;++s)
      for(int h0=0;h0<H;++h0){
        int tau=s*H+h0; double hrs=Hours[idx_hours(s,h0,H)];

        // Demand row: sum Gen + inflows - outflows = Demand
        int rowD=ROW_DEM0+r*T+tau;
        for(int e0=0;e0<E;++e0){
          int c=base+col_gen_e_tau(E,e0,tau,T); Tpls.push_back({rowD,c,1.0});
        }
        for(int d=0; d<R; ++d){ if(d==r) continue;
          int base_d=d*g_colsPerReg; int cIn=base_d+col_flow_od_tau(E,T,R,d,r,tau);
          Tpls.push_back({rowD,cIn,1.0});
        }
        for(int d=0; d<R; ++d){ if(d==r) continue;
          int cOut=base+col_flow_od_tau(E,T,R,r,d,tau); Tpls.push_back({rowD,cOut,-1.0});
        }

        // Capacity rows per (r,e)
        for(int e0=0;e0<E;++e0){
          double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
          int rowC=ROW_CAP0+((r*E+e0)*T+tau);
          int cG=base+col_gen_e_tau(E,e0,tau,T);
          int cB=base+col_build_e(e0);
          Tpls.push_back({rowC,cG,-1.0});
          if(cf!=0.0 && hrs!=0.0) Tpls.push_back({rowC,cB,cf*hrs});
        }

        // Reserve rows: only local capacity contributes
        if(useRes){
          int rowR=ROW_RSV0+r*T+tau;
          for(int e0=0;e0<E;++e0){
            double cf=CF[idx_cf(r,e0,s,h0,E,S,H)];
            if(cf!=0.0 && hrs!=0.0){
              int cB=base+col_build_e(e0); Tpls.push_back({rowR,cB,cf*hrs});
            }
          }
        }

        // CES row
        if(useCES){
          int rowCES=ROW_CES;
          for(int e0=0;e0<E;++e0){
            double q=CES_q[e0];
            if(q!=0.0){
              int cG=base+col_gen_e_tau(E,e0,tau,T);
              Tpls.push_back({rowCES,cG,q});
            }
          }
        }

        // RPS row
        if(useRPS){
          int rowRPS=ROW_RPS;
          for(int e0=0;e0<E;++e0){
            double q=RPS_q[e0];
            if(q!=0.0){
              int cG=base+col_gen_e_tau(E,e0,tau,T);
              Tpls.push_back({rowRPS,cG,q});
            }
          }
        }
      }
  }

  // Slack coefficients
  if(useCES && g_colSlackCES>=0) Tpls.push_back({ROW_CES,g_colSlackCES,1.0});
  if(useRPS && g_colSlackRPS>=0) Tpls.push_back({ROW_RPS,g_colSlackRPS,1.0});

  // Convert to CSC for HiGHS
  vector<int> astart_i,aindex_i; vector<double> avalue;
  triplets_to_csc(rowCount,g_N_total,Tpls,astart_i,aindex_i,avalue);
  vector<HighsInt> astart(astart_i.begin(),astart_i.end());
  vector<HighsInt> aindex(aindex_i.begin(),aindex_i.end());
  HighsInt nnz=(HighsInt)avalue.size();

  // Solve
  void* h=Highs_create();
  const HighsInt a_format=1, sense=1; const double obj_offset=0.0;
  g_lastCode=Highs_passLp(
    h,(HighsInt)g_N_total,(HighsInt)rowCount,nnz,
    a_format,sense,obj_offset,
    g_N_total? &col_cost[0]:nullptr, g_N_total? &col_lo[0]:nullptr, g_N_total? &col_hi[0]:nullptr,
    rowCount? &row_lo[0]:nullptr,    rowCount? &row_hi[0]:nullptr,
    astart.data(), aindex.data(), avalue.data()
  );
  if(g_lastCode==0) g_lastCode=Highs_run(h);
  g_status = (g_lastCode==0)? Highs_getModelStatus(h) : -1;
  g_obj=1e308; g_solved=false;

  if(g_lastCode==0){
    double obj=Highs_getObjectiveValue(h);
    vector<double> col_val(g_N_total,0.0), col_dual(g_N_total,0.0);
    vector<double> row_val(std::max(1,rowCount),0.0), row_dual(std::max(1,rowCount),0.0);
    if(Highs_getSolution(h,col_val.data(),col_dual.data(),row_val.data(),row_dual.data())==0){
      g_solution=std::move(col_val);
      // Region cost breakout
      for(int r=0;r<g_R;++r){
        double cc=0.0,ce=0.0; int base=r*g_colsPerReg;
        for(int e0=0;e0<g_E;++e0){
          int cB=base+col_build_e(e0);
          cc += (CAPEX[e0]+FOM[e0])*g_solution[cB];
          double vc=Fuel[e0]+VOM[e0];
          for(int tau=0;tau<g_T;++tau){
            int cG=base+col_gen_e_tau(g_E,e0,tau,g_T);
            ce += vc*g_solution[cG];
          }
        }
        g_cost_cap[r]=cc; g_cost_energy[r]=ce;
      }
      g_obj=obj; g_solved=std::isfinite(g_obj);
    }
  }
  Highs_destroy(h);
  return g_solved? g_obj : 1e308;
}

// ---------- Vensim glue ----------
static const int EXTERN_VCODE=62051;
extern "C" __declspec(dllexport) int VEFCC version_info(){ return EXTERN_VCODE; }

enum {
  F_SOLVE=1101,   // LP_Solve(..., scenario_id)
  F_CAP_ADD,      // LP_CAP_ADD(r,e, scenario_id)
  F_GEN,          // LP_GEN(r,e,ts,hr, scenario_id)
  F_COST_CAP,     // LP_COST_CAP(r, scenario_id)
  F_COST_ENERGY,  // LP_COST_ENERGY(r, scenario_id)
  F_LP_OBJ,       // LP_OBJ(scenario_id)
  F_LP_STATUS,    // LP_STATUS(scenario_id)
  F_LP_CODE,      // LP_CODE(scenario_id)
  F_FLOW_EXPORT,  // LP_EXPORT(r,ts,hr, scenario_id)
  F_FLOW_IMPORT,  // LP_IMPORT(r,ts,hr, scenario_id)
  F_FLOW_OD       // LP_FLOW_OD(o,d,ts,hr, scenario_id)
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
      // 13 vectors + scalars (CES_rhs, CES_ACP, RPS_rhs, RPS_ACP, ReserveMargin, R,S,H,E, scenario_id)
      *arglist=(char*)"{Demand},{Capacity_factor},{Hours},{CAPEX},{FOM},{Fuel},{VOM},{ExistingCap},{MaxBuild},{CapMax},{TransRR},{CES_qualifying},{RPS_qualifying},CES_rhs,CES_ACP,RPS_rhs,RPS_ACP,ReserveMargin,R,S,H,E,scenario_id";
      *num_args=23; *num_vector=13; *func_index=F_SOLVE; return 1;

    case 1:  *sym=(char*)"LP_CAP_ADD";       *arglist=(char*)" region_pos , tech_pos , scenario_id ";                  *num_args=3; *num_vector=0; *func_index=F_CAP_ADD;      return 1;
    case 2:  *sym=(char*)"LP_GEN";           *arglist=(char*)" region_pos , tech_pos , ts_pos , hr_pos , scenario_id "; *num_args=5; *num_vector=0; *func_index=F_GEN;          return 1;
    case 3:  *sym=(char*)"LP_COST_CAP";      *arglist=(char*)" region_pos , scenario_id ";                              *num_args=2; *num_vector=0; *func_index=F_COST_CAP;     return 1;
    case 4:  *sym=(char*)"LP_COST_ENERGY";   *arglist=(char*)" region_pos , scenario_id ";                              *num_args=2; *num_vector=0; *func_index=F_COST_ENERGY;  return 1;
    case 5:  *sym=(char*)"LP_OBJ";           *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_OBJ;       return 1;
    case 6:  *sym=(char*)"LP_STATUS";        *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_STATUS;    return 1;
    case 7:  *sym=(char*)"LP_CODE";          *arglist=(char*)" scenario_id ";                                          *num_args=1; *num_vector=0; *func_index=F_LP_CODE;      return 1;
    case 8:  *sym=(char*)"LP_EXPORT";        *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_FLOW_EXPORT;  return 1;
    case 9:  *sym=(char*)"LP_IMPORT";        *arglist=(char*)" region_pos , ts_pos , hr_pos , scenario_id ";           *num_args=4; *num_vector=0; *func_index=F_FLOW_IMPORT;  return 1;
    case 10: *sym=(char*)"LP_FLOW_OD";       *arglist=(char*)" origin_pos , dest_pos , ts_pos , hr_pos , scenario_id ";*num_args=5; *num_vector=0; *func_index=F_FLOW_OD;     return 1;
    default: return 0;
  }
}

extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch(funcid){

    case F_SOLVE: {
      if(nval<22){ val[0].val=1e308; return 0; }

      const double* Dem  = val[0].vec->firstval;
      const double* CF   = val[1].vec->firstval;
      const double* Hours= val[2].vec->firstval;
      const double* CAPEX= val[3].vec->firstval;
      const double* FOM  = val[4].vec->firstval;
      const double* Fuel = val[5].vec->firstval;
      const double* VOM  = val[6].vec->firstval;
      const double* XCap = val[7].vec->firstval;
      const double* MaxB = val[8].vec->firstval;
      const double* CapMx= val[9].vec->firstval;
      const double* Trans= val[10].vec->firstval;
      const double* CES_q= val[11].vec->firstval;
      const double* RPS_q= val[12].vec->firstval;

      double CES_rhs = val[13].val;
      double CES_ACP = val[14].val;
      double RPS_rhs = val[15].val;
      double RPS_ACP = val[16].val;

      double ReserveMargin = val[17].val;
      int R=(int)(val[18].val+0.5), S=(int)(val[19].val+0.5), H=(int)(val[20].val+0.5), E=(int)(val[21].val+0.5);

      int scenario_id = (nval>=23)? (int)(val[22].val+0.5) : 0;

      val[0].val = solve_global_policies(
        Dem,CF,Hours,CAPEX,FOM,Fuel,VOM,XCap,MaxB,CapMx,Trans,
        CES_q,CES_rhs,CES_ACP, RPS_q,RPS_rhs,RPS_ACP,
        ReserveMargin,R,S,H,E
      );
      save_context(scenario_id);
      return 1;
    }

    case F_CAP_ADD: {
      if(nval<3){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), sid=(int)(val[2].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; val[0].val=c->solution[base+col_build_e(e-1)];
      return 1;
    }

    case F_GEN: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), e=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||e<1||e>c->E||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int base=(r-1)*c->colsPerReg; int tau=(ts-1)*c->H+(hr-1);
      val[0].val=c->solution[base+col_gen_e_tau(c->E,e-1,tau,c->T)];
      return 1;
    }

    case F_COST_CAP: {
      if(nval<2){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R){ val[0].val=1e308; return 1; }
      val[0].val=c->cost_cap[r-1]; return 1;
    }

    case F_COST_ENERGY: {
      if(nval<2){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), sid=(int)(val[1].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R){ val[0].val=1e308; return 1; }
      val[0].val=c->cost_energy[r-1]; return 1;
    }

    case F_LP_OBJ: {
      if(nval<1){ val[0].val=1e308; return 1; }
      int sid=(int)(val[0].val+0.5);
      const Context* c=get_ctx(sid);
      val[0].val = c? c->obj : 1e308;
      return 1;
    }

    case F_LP_STATUS: {
      if(nval<1){ val[0].val=1e308; return 1; }
      int sid=(int)(val[0].val+0.5);
      const Context* c=get_ctx(sid);
      val[0].val = (double)(c? c->status : -1);
      return 1;
    }

    case F_LP_CODE: {
      if(nval<1){ val[0].val=1e308; return 1; }
      int sid=(int)(val[0].val+0.5);
      const Context* c=get_ctx(sid);
      val[0].val = (double)(c? c->lastCode : -999);
      return 1;
    }

    case F_FLOW_EXPORT: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int r0=r-1, tau=(ts-1)*c->H+(hr-1);
      double sum=0.0;
      int base_r = r0 * c->colsPerReg;
      for(int d0=0; d0<c->R; ++d0){
        if(d0==r0) continue;
        int col = base_r + col_flow_od_tau(c->E,c->T,c->R, r0,d0,tau);
        sum += c->solution[col];
      }
      val[0].val = sum; return 1;
    }

    case F_FLOW_IMPORT: {
      if(nval<4){ val[0].val=1e308; return 1; }
      int r=(int)(val[0].val+0.5), ts=(int)(val[1].val+0.5), hr=(int)(val[2].val+0.5), sid=(int)(val[3].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||r<1||r>c->R||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int r0=r-1, tau=(ts-1)*c->H+(hr-1);
      double sum=0.0;
      for(int o0=0; o0<c->R; ++o0){
        if(o0==r0) continue;
        int base_o = o0 * c->colsPerReg;
        int col = base_o + col_flow_od_tau(c->E,c->T,c->R, o0,r0,tau);
        sum += c->solution[col];
      }
      val[0].val = sum; return 1;
    }

    case F_FLOW_OD: {
      if(nval<5){ val[0].val=1e308; return 1; }
      int o=(int)(val[0].val+0.5), d=(int)(val[1].val+0.5), ts=(int)(val[2].val+0.5), hr=(int)(val[3].val+0.5), sid=(int)(val[4].val+0.5);
      const Context* c=get_ctx(sid);
      if(!c||!c->solved||o<1||o>c->R||d<1||d>c->R||o==d||ts<1||ts>c->S||hr<1||hr>c->H){ val[0].val=1e308; return 1; }
      int o0=o-1, d0=d-1, tau=(ts-1)*c->H+(hr-1);
      int base_o = o0 * c->colsPerReg;
      int col = base_o + col_flow_od_tau(c->E,c->T,c->R, o0,d0,tau);
      val[0].val = c->solution[col]; return 1;
    }

    default: return 0;
  }
}
