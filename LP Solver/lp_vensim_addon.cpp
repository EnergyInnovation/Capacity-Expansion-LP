// lp_vensim_addon.cpp — HiGHS C API (15-arg Highs_passLp)
// Region-first layout; global CES; per-hour reserve; per-region cost breakout.
//
// Build (x64 Native Tools for VS):
//   del /q lp_vensim_addon.obj lp_vensim_addon.lib lp_vensim_addon.exp lp_vensim_addon.dll
//   cl /nologo /LD /O2 /EHsc /MT ^
//     "lp_vensim_addon.cpp" ^
//     /I "C:\vcpkg\installed\x64-windows-static\include" ^
//     /I "C:\vcpkg\installed\x64-windows-static\include\highs" ^
//     /link /NOLOGO ^
//     /LIBPATH:"C:\vcpkg\installed\x64-windows-static\lib" highs.lib ^
//     /OUT:"lp_vensim_addon.dll"
// If you get LNK2038 CRT mismatch, switch /MT -> /MD to match highs.lib.

#define NOMINMAX
#include <windows.h>
#include <vector>
#include <mutex>
#include <algorithm>
#include <cstdint>
#include <cmath>

#include <highs/interfaces/highs_c_api.h>  // vcpkg header

// ===== Vensim ABI =====
#if defined(_MSC_VER)
  #define VEFCC __stdcall
#else
  #define VEFCC
#endif
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

// ---------- Global cache ----------
static std::mutex          g_mutex;
static std::vector<double> g_solution;    // size N
static std::vector<double> g_cost_cap;    // size R
static std::vector<double> g_cost_energy; // size R
static int                 g_status   = -1;   // HiGHS model status
static int                 g_lastCode = -999; // HiGHS API return code
static double              g_obj      = 0.0;
static bool                g_solved   = false;

// Shape of last solve
static int g_R=0, g_S=0, g_H=0, g_E=0;
static int g_T=0, g_colsPerReg=0, g_N=0;

static inline void reset_cache(int R=0,int E=0,int T=0){
  int colsPerReg = (E>0 ? E*(T+1) : 0);
  int N = (R>0 ? R*colsPerReg : 0);
  g_solution.assign(N, 0.0);
  g_cost_cap.assign(R>0?R:0, 0.0);
  g_cost_energy.assign(R>0?R:0, 0.0);
  g_status=-1; g_lastCode=-999; g_obj=0.0; g_solved=false;
  g_R=R; g_E=E; g_T=T; g_colsPerReg=colsPerReg; g_N=N;
}

// ---------- Index helpers (Region-first) ----------
static inline int idx_dem(int r,int s,int h,int S,int H){ return ((r*S)+s)*H + h; }                    // [R,S,H]
static inline int idx_hours(int s,int h,int H){ return s*H + h; }                                       // [S,H]
static inline int idx_cf(int r,int e,int s,int h,int E,int S,int H){ return (((r*E + e)*S + s)*H + h); } // [R,E,S,H]
static inline int idx_xcap(int r,int e,int E){ return r*E + e; }                                        // [R,E]

// Columns (0-based within region block)
static inline int col_build_e(int e) { return e; } // 0..E-1
static inline int col_gen_e_tau(int E, int e, int tau, int T){ return E + e*T + tau; } // 0..E*T-1

// ---------- Triplet -> CSC ----------
struct Triplet { int row; int col; double val; };

static void triplets_to_csc(
  int nRows, int nCols,
  const std::vector<Triplet>& t,
  std::vector<int>& astart,
  std::vector<int>& aindex,
  std::vector<double>& avalue
){
  astart.assign(nCols+1, 0);
  for (const auto& x : t) { if (x.col>=0 && x.col<nCols) astart[x.col+1]++; }
  for (int c=0; c<nCols; ++c) astart[c+1] += astart[c];
  int nnz = (int)t.size();
  aindex.assign(nnz, 0);
  avalue.assign(nnz, 0.0);
  std::vector<int> next = astart;
  for (const auto& x : t){
    int p = next[x.col]++;
    aindex[p] = x.row;
    avalue[p] = x.val;
  }
}

// ---------- Core solver (HiGHS C API: 15-arg Highs_passLp) ----------
static double solve_global_ces_reserve(
  // Data (Region-first)
  const double* Demand,            // [R,S,H]
  const double* CF,                // [R,E,S,H]
  const double* Hours,             // [S,H]
  const double* CAPEX,             // [E]
  const double* FOM,               // [E]
  const double* Fuel,              // [E]
  const double* VOM,               // [E]
  const double* ExistingCap,       // [R,E]
  const double* CES_q,             // [E] weights
  double CES_rhs,                  // scalar
  double ReserveMargin,            // scalar (e.g., 0.15)
  int R, int S, int H, int E
){
  const int T = S*H;
  reset_cache(R,E,T);

  // Row indexing
  const int ROW_DEM0   = 0;               // R*T
  const int ROW_CAP0   = ROW_DEM0 + R*T;  // R*E*T
  const int ROW_RSV0   = ROW_CAP0 + R*E*T;// R*T if used
  int rowCount         = ROW_CAP0 + R*E*T;
  const bool useReserve = (ReserveMargin > 0.0);
  if (useReserve) rowCount += R*T;

  bool useCES = (CES_rhs > 0.0);
  if (useCES){
    bool anyQ=false; for (int e=0;e<E;++e) if (CES_q[e]!=0.0){ anyQ=true; break; }
    if (!anyQ) useCES=false;
  }
  const int ROW_CES = rowCount;
  if (useCES) rowCount += 1;

  // Columns
  const double INF = 1e30;
  std::vector<double> col_cost(g_N, 0.0), col_lo(g_N, 0.0), col_hi(g_N, INF);

  for (int r=0; r<R; ++r){
    int base = r * g_colsPerReg;
    // Build vars
    for (int e=0; e<E; ++e){
      int col = base + col_build_e(e);
      col_cost[col] = CAPEX[e] + FOM[e];
      col_lo[col]   = 0.0; col_hi[col] = INF;
    }
    // Gen vars
    for (int e=0; e<E; ++e){
      double c_e = Fuel[e] + VOM[e];
      for (int tau=0; tau<T; ++tau){
        int col = base + col_gen_e_tau(E, e, tau, T);
        col_cost[col] = c_e;
        col_lo[col]   = 0.0; col_hi[col] = INF;
      }
    }
  }

  // Row bounds
  std::vector<double> row_lo(rowCount, -INF), row_hi(rowCount, INF);

  // Demand (=)
  for (int r=0; r<R; ++r)
    for (int s=0; s<S; ++s)
      for (int h0=0; h0<H; ++h0){
        int tau = s*H + h0;
        int row = ROW_DEM0 + r*T + tau;
        double rhs = Demand[idx_dem(r,s,h0,S,H)];
        row_lo[row] = rhs; row_hi[row] = rhs;
      }

  // Capacity (≥)
  for (int r=0; r<R; ++r)
    for (int e=0; e<E; ++e){
      double xcap = ExistingCap[idx_xcap(r,e,E)];
      for (int s=0; s<S; ++s)
        for (int h0=0; h0<H; ++h0){
          int tau = s*H + h0;
          double cf  = CF[idx_cf(r,e,s,h0,E,S,H)];
          double hrs = Hours[idx_hours(s,h0,H)];
          int row = ROW_CAP0 + ((r*E + e)*T + tau);
          row_lo[row] = -cf * hrs * xcap;  // lower bound
          row_hi[row] = INF;
        }
    }

  // Reserve (≥)
  if (useReserve){
    for (int r=0; r<R; ++r)
      for (int s=0; s<S; ++s)
        for (int h0=0; h0<H; ++h0){
          int tau = s*H + h0;
          double rhs = (1.0 + ReserveMargin) * Demand[idx_dem(r,s,h0,S,H)];
          double hrs = Hours[idx_hours(s,h0,H)];
          for (int e=0; e<E; ++e){
            rhs -= CF[idx_cf(r,e,s,h0,E,S,H)] * hrs * ExistingCap[idx_xcap(r,e,E)];
          }
          int row = ROW_RSV0 + r*T + tau;
          row_lo[row] = rhs; row_hi[row] = INF;
        }
  }

  // CES (≥)
  if (useCES){
    row_lo[ROW_CES] = CES_rhs; row_hi[ROW_CES] = INF;
  }

  // Triplets
  std::vector<Triplet> Tpls;
  Tpls.reserve(
    (size_t)R*(size_t)E*(size_t)T*3
    + (useReserve? (size_t)R*(size_t)T*(size_t)E : 0)
    + (useCES? (size_t)R*(size_t)E*(size_t)T : 0)
    + (size_t)R*(size_t)T*(size_t)E
  );

  for (int r=0; r<R; ++r){
    int base = r * g_colsPerReg;

    for (int s=0; s<S; ++s)
      for (int h0=0; h0<H; ++h0){
        int tau = s*H + h0;
        double hrs = Hours[idx_hours(s,h0,H)];

        // Demand row
        int rowD = ROW_DEM0 + r*T + tau;
        for (int e=0; e<E; ++e){
          int colG = base + col_gen_e_tau(E, e, tau, T);
          Tpls.push_back({rowD, colG, 1.0});
        }

        // Capacity rows
        for (int e=0; e<E; ++e){
          double cf = CF[idx_cf(r,e,s,h0,E,S,H)];
          int rowC  = ROW_CAP0 + ((r*E + e)*T + tau);
          int colG  = base + col_gen_e_tau(E, e, tau, T);
          int colB  = base + col_build_e(e);
          Tpls.push_back({rowC, colG, -1.0});
          if (cf!=0.0 && hrs!=0.0)
            Tpls.push_back({rowC, colB, cf * hrs});
        }

        // Reserve rows
        if (useReserve){
          int rowR = ROW_RSV0 + r*T + tau;
          for (int e=0; e<E; ++e){
            double cf = CF[idx_cf(r,e,s,h0,E,S,H)];
            if (cf!=0.0 && hrs!=0.0){
              int colB = base + col_build_e(e);
              Tpls.push_back({rowR, colB, cf * hrs});
            }
          }
        }

        // CES row
        if (useCES){
          int rowC = ROW_CES;
          for (int e=0; e<E; ++e){
            double q = CES_q[e];
            if (q!=0.0){
              int colG = base + col_gen_e_tau(E, e, tau, T);
              Tpls.push_back({rowC, colG, q});
            }
          }
        }
      }
  }

  // Convert to CSC for HiGHS
  std::vector<int> astart_i; std::vector<int> aindex_i; std::vector<double> avalue;
  triplets_to_csc(rowCount, g_N, Tpls, astart_i, aindex_i, avalue);

  // HiGHS expects HighsInt for starts/indices
  std::vector<HighsInt> astart(astart_i.begin(), astart_i.end());
  std::vector<HighsInt> aindex(aindex_i.begin(), aindex_i.end());
  const HighsInt nnz = (HighsInt)avalue.size();

  // ---- HiGHS solve ----
  void* h = Highs_create();

  // 15-arg signature:
  // Highs_passLp(h, num_col, num_row, num_nz, a_format, sense, offset,
  //   col_cost, col_lower, col_upper, row_lower, row_upper,
  //   a_start, a_index, a_value);
  const HighsInt a_format = 1;  // 1 = column-wise (CSC)
  const HighsInt sense    = 1;  // 1 = minimize
  const double   offset   = 0.0;

  g_lastCode = Highs_passLp(
    h,
    (HighsInt)g_N, (HighsInt)rowCount, nnz,
    a_format, sense, offset,
    col_cost.data(), col_lo.data(), col_hi.data(),
    row_lo.data(), row_hi.data(),
    astart.data(), aindex.data(), avalue.data()
  );

  if (g_lastCode==0) g_lastCode = Highs_run(h);  // 0 = ok

  int ms = -1;
  if (g_lastCode==0) ms = Highs_getModelStatus(h);

  g_status = ms;
  g_obj    = 1e308;
  g_solved = false;

  if (g_lastCode==0){
    double obj = Highs_getObjectiveValue(h);

    std::vector<double> col_val(g_N, 0.0), col_dual(g_N,0.0);
    std::vector<double> row_val(std::max(1,rowCount), 0.0), row_dual(std::max(1,rowCount), 0.0);
    int sol_rc = Highs_getSolution(h, col_val.data(), col_dual.data(), row_val.data(), row_dual.data());
    if (sol_rc==0) {
      g_solution = std::move(col_val);
      // Cost breakout (by region)
      for (int r=0; r<g_R; ++r){
        double cc=0.0, ce=0.0;
        int base = r * g_colsPerReg;
        for (int e=0; e<g_E; ++e){
          int colB = base + col_build_e(e);
          cc += (CAPEX[e] + FOM[e]) * g_solution[colB];
          double varc = Fuel[e] + VOM[e];
          for (int tau=0; tau<g_T; ++tau){
            int colG = base + col_gen_e_tau(g_E, e, tau, g_T);
            ce += varc * g_solution[colG];
          }
        }
        if ((int)g_cost_cap.size()!=g_R) g_cost_cap.assign(g_R,0.0);
        if ((int)g_cost_energy.size()!=g_R) g_cost_energy.assign(g_R,0.0);
        g_cost_cap[r]    = cc;
        g_cost_energy[r] = ce;
      }
      g_obj    = obj;
      g_solved = std::isfinite(g_obj);
    }
  }

  Highs_destroy(h);
  return g_solved ? g_obj : 1e308;
}

// ---------- Vensim glue ----------
static const int EXTERN_VCODE = 62051;
extern "C" __declspec(dllexport) int VEFCC version_info(){ return EXTERN_VCODE; }

enum {
  F_SOLVE = 1101,
  F_CAP_ADD,
  F_GEN,
  F_COST_CAP,
  F_COST_EN,
  F_LP_OBJ,
  F_LP_STATUS,
  F_LP_CODE
};

extern "C" __declspec(dllexport) int VEFCC user_definition(
  int setup_index,
  char** sym, char** arglist,
  int* num_args, int* num_vector,
  int* func_index, int* dim_act,
  int* modify, int* num_loops,
  int* num_literal, int* num_lookup
){
  if (!sym||!arglist||!num_args||!num_vector||!func_index||!dim_act||!modify||!num_loops||!num_literal||!num_lookup) return 0;
  *dim_act=0; *modify=0; *num_loops=0; *num_literal=0; *num_lookup=0;

  switch (setup_index){
    case 0:
      *sym        = (char*)"LP_SOLVE_RTH_XCAP_CES_RES";
      *arglist    = (char*)"{Demand},{Capacity_factor},{Hours},{CAPEX},{FOM},{Fuel},{VOM},{ExistingCap},{CES_qualifying},CES_rhs,ReserveMargin,R,S,H,E";
      *num_args   = 15;
      *num_vector = 9;     // first 9 are vectors
      *func_index = F_SOLVE;
      return 1;

    case 1: *sym=(char*)"LP_CAP_ADD";     *arglist=(char*)" region_pos , tech_pos ";                  *num_args=2; *num_vector=0; *func_index=F_CAP_ADD; return 1;
    case 2: *sym=(char*)"LP_GEN";         *arglist=(char*)" region_pos , tech_pos , ts_pos , hr_pos "; *num_args=4; *num_vector=0; *func_index=F_GEN;     return 1;
    case 3: *sym=(char*)"LP_COST_CAP";    *arglist=(char*)" region_pos ";                              *num_args=1; *num_vector=0; *func_index=F_COST_CAP; return 1;
    case 4: *sym=(char*)"LP_COST_ENERGY"; *arglist=(char*)" region_pos ";                              *num_args=1; *num_vector=0; *func_index=F_COST_EN;  return 1;

    case 5: *sym=(char*)"LP_OBJ";         *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_OBJ;    return 1;
    case 6: *sym=(char*)"LP_STATUS";      *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_STATUS; return 1;
    case 7: *sym=(char*)"LP_CODE";        *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_CODE;   return 1;

    default: return 0;
  }
}

extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch (funcid){

    case F_SOLVE: {
      if (nval < 15) { val[0].val = 1e308; return 0; }

      const double* Dem    = val[0].vec->firstval; // [R,S,H]
      const double* CF     = val[1].vec->firstval; // [R,E,S,H]
      const double* Hours  = val[2].vec->firstval; // [S,H]
      const double* CAPEX  = val[3].vec->firstval; // [E]
      const double* FOM    = val[4].vec->firstval; // [E]
      const double* Fuel   = val[5].vec->firstval; // [E]
      const double* VOM    = val[6].vec->firstval; // [E]
      const double* XCap   = val[7].vec->firstval; // [R,E]
      const double* CES_q  = val[8].vec->firstval; // [E]

      double CES_rhs       = val[9].val;
      double ReserveMargin = val[10].val;
      int R = (int)(val[11].val + 0.5);
      int S = (int)(val[12].val + 0.5);
      int H = (int)(val[13].val + 0.5);
      int E = (int)(val[14].val + 0.5);

      val[0].val = solve_global_ces_reserve(
        Dem, CF, Hours, CAPEX, FOM, Fuel, VOM, XCap, CES_q,
        CES_rhs, ReserveMargin, R,S,H,E
      );
      return 1;
    }

    case F_CAP_ADD: {
      if (!g_solved || nval<2){ val[0].val=1e308; return 1; }
      int r = (int)(val[0].val + 0.5);
      int e = (int)(val[1].val + 0.5);
      if (r<1||r>g_R||e<1||e>g_E){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      int col  = base + col_build_e(e-1);
      val[0].val = g_solution[col];
      return 1;
    }

    case F_GEN: {
      if (!g_solved || nval<4){ val[0].val=1e308; return 1; }
      int r  = (int)(val[0].val + 0.5);
      int e  = (int)(val[1].val + 0.5);
      int ts = (int)(val[2].val + 0.5);
      int hr = (int)(val[3].val + 0.5);
      if (r<1||r>g_R||e<1||e>g_E||ts<1||ts>g_S||hr<1||hr>g_H){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      int tau  = (ts-1)*g_H + (hr-1);
      int col  = base + col_gen_e_tau(g_E, e-1, tau, g_T);
      val[0].val = g_solution[col];
      return 1;
    }

    case F_COST_CAP: {
      if (!g_solved || nval<1){ val[0].val=1e308; return 1; }
      int r = (int)(val[0].val + 0.5);
      if (r<1||r>g_R){ val[0].val=1e308; return 1; }
      val[0].val = g_cost_cap[r-1];
      return 1;
    }

    case F_COST_EN: {
      if (!g_solved || nval<1){ val[0].val=1e308; return 1; }
      int r = (int)(val[0].val + 0.5);
      if (r<1||r>g_R){ val[0].val=1e308; return 1; }
      val[0].val = g_cost_energy[r-1];
      return 1;
    }

    case F_LP_OBJ:    { val[0].val = g_obj;       return 1; }
    case F_LP_STATUS: { val[0].val = (double)g_status;   return 1; }
    case F_LP_CODE:   { val[0].val = (double)g_lastCode; return 1; }

    default: return 0;
  }
}
