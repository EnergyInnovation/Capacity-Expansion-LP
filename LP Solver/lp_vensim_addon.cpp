// lp_vensim_addon.cpp
// GLPK-based LP addon for Vensim (single solver; global CES; hourly reserve margin).
//
// Column layout per region r (0..R-1):
//   [Build[r, e=0..E-1]]  (E vars)
//   [Gen[r, e=0..E-1, tau=0..T-1]]  (E*T vars, with tau = s*H + h)
// Total per region: E + E*T = E*(T+1).  Total N = R * E*(T+1).
//
// Rows:
//  (1) Demand balance (R*T rows):   Sum_e Gen[r,e,tau] = Demand[r,tau]
//  (2) Capacity limit (R*E*T rows): -Gen[r,e,tau] + CF[r,e,tau]*Hours[tau]*Build[r,e] >= -CF[r,e,tau]*Hours[tau]*Existing[r,e]
//  (3) Reserve (R*T rows, optional if RM>0):
//        Sum_e CF[r,e,tau]*Hours[tau]*Build[r,e] >= (1+RM)*Demand[r,tau] - Sum_e CF[r,e,tau]*Hours[tau]*Existing[r,e]
//  (4) Global CES (1 row, optional if CES_rhs>0 && any q[e]!=0):
//        Sum_{r,e,tau} q[e] * Gen[r,e,tau] >= CES_rhs
//
// Objective:
//   Minimize Sum_r [ Sum_e (CAPEX[e]+FOM[e])*Build[r,e] + Sum_{e,tau} (Fuel[e]+VOM[e])*Gen[r,e,tau] ]
//
// Exports:
//   LP_SOLVE_RTH_XCAP_CES_RES(
//     {Demand},{Capacity_factor},{Hours},
//     {CAPEX},{FOM},{Fuel},{VOM},
//     {ExistingCap},{CES_qualifying},
//     CES_rhs, ReserveMargin, R,S,H,E)
//   Getters:
//     LP_CAP_ADD(r_pos, e_pos)
//     LP_GEN(r_pos, e_pos, ts_pos, hr_pos)
//     LP_COST_CAP(r_pos)   -- capacity cost for region r
//     LP_COST_ENERGY(r_pos)-- energy (variable) cost for region r
//     LP_OBJ(), LP_STATUS(), LP_CODE()
//
// Build (VS 2022 x64, static GLPK):
//   cl /nologo /LD /O2 /EHsc /MT ^
//     "C:\\Users\\RobbieOrvis\\Models\\Capacity Expansion LP\\LP Solver\\lp_vensim_addon.cpp" ^
//     /I "C:\\vcpkg\\installed\\x64-windows-static\\include" ^
//     /link /NOLOGO ^
//     /LIBPATH:"C:\\vcpkg\\installed\\x64-windows-static\\lib" glpk.lib ^
//     /OUT:"C:\\Users\\RobbieOrvis\\Models\\Capacity Expansion LP\\LP Solver\\lp_vensim_addon.dll"

#define NOMINMAX
#include <windows.h>
#include <vector>
#include <mutex>
#include <algorithm>

extern "C" {
  typedef struct glp_prob glp_prob;
  enum { GLP_MIN=1 };
  enum { GLP_FR=1, GLP_LO=2, GLP_UP=3, GLP_DB=4, GLP_FX=5 };
  enum { GLP_UNDEF=1, GLP_FEAS=2, GLP_INFEAS=3, GLP_NOFEAS=4, GLP_OPT=5, GLP_UNBND=6 };

  glp_prob* glp_create_prob(void);
  void      glp_set_prob_name(glp_prob*, const char*);
  void      glp_set_obj_dir(glp_prob*, int);
  void      glp_add_rows(glp_prob*, int);
  void      glp_set_row_bnds(glp_prob*, int, int, double, double);
  void      glp_add_cols(glp_prob*, int);
  void      glp_set_obj_coef(glp_prob*, int, double);
  void      glp_set_col_bnds(glp_prob*, int, int, double, double);
  void      glp_load_matrix(glp_prob*, int, const int[], const int[], const double[]);
  int       glp_simplex(glp_prob*, const void*);
  int       glp_get_status(glp_prob*);
  double    glp_get_obj_val(glp_prob*);
  double    glp_get_col_prim(glp_prob*, int);
  void      glp_delete_prob(glp_prob*);
}

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
static std::vector<double> g_solution;          // size N
static std::vector<double> g_cost_cap;          // size R
static std::vector<double> g_cost_energy;       // size R
static int                 g_status   = -1;
static int                 g_lastCode = -999;
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

// ---------- Index helpers (Region-first order) ----------
static inline int idx_dem(int r,int s,int h,int S,int H){ return ((r*S)+s)*H + h; }                 // [R,S,H]
static inline int idx_hours(int s,int h,int H){ return s*H + h; }                                    // [S,H]
static inline int idx_cf(int r,int e,int s,int h,int E,int S,int H){ return (((r*E + e)*S + s)*H + h); } // [R,E,S,H]
static inline int idx_xcap(int r,int e,int E){ return r*E + e; }                                     // [R,E]

// Columns (0-based within region block)
static inline int col_build_e(int e) { return e; } // 0..E-1
static inline int col_gen_e_tau(int E, int e, int tau, int T){ return E + e*T + tau; } // 0..E*T-1

// ---------- Single solver ----------
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

  glp_prob* lp = glp_create_prob();
  glp_set_prob_name(lp, "vensim_lp_global_ces_reserve");
  glp_set_obj_dir(lp, GLP_MIN);

  // Row indexing segments
  const int ROW_DEM0   = 0;               // size R*T
  const int ROW_CAP0   = ROW_DEM0 + R*T;  // size R*E*T
  const int ROW_RSV0   = ROW_CAP0 + R*E*T;// size R*T (optional if RM>0)
  int rowCount         = ROW_CAP0 + R*E*T;
  bool useReserve      = (ReserveMargin > 0.0);
  if (useReserve) rowCount += R*T;

  // CES row (optional)
  bool useCES = (CES_rhs > 0.0);
  if (useCES){
    bool anyQ = false;
    for (int e=0; e<E; ++e) if (CES_q[e] != 0.0){ anyQ = true; break; }
    if (!anyQ) useCES = false;
  }
  const int ROW_CES = rowCount; // if used, 1 extra row at end
  if (useCES) rowCount += 1;

  glp_add_rows(lp, rowCount);

  // (1) Demand rows (=)
  for (int r=0; r<R; ++r){
    for (int s=0; s<S; ++s){
      for (int h=0; h<H; ++h){
        int tau = s*H + h;
        int row = ROW_DEM0 + r*T + tau + 1;
        double rhs = Demand[idx_dem(r,s,h,S,H)];
        glp_set_row_bnds(lp, row, GLP_FX, rhs, rhs);
      }
    }
  }

  // (2) Capacity rows (>= -CF*hrs*Existing)
  for (int r=0; r<R; ++r){
    for (int e=0; e<E; ++e){
      double xcap = ExistingCap[idx_xcap(r,e,E)];
      for (int s=0; s<S; ++s){
        for (int h=0; h<H; ++h){
          int tau = s*H + h;
          double cf   = CF[idx_cf(r,e,s,h,E,S,H)];
          double hrs  = Hours[idx_hours(s,h,H)];
          double lb   = -cf * hrs * xcap;
          int row     = ROW_CAP0 + ( (r*E + e)*T + tau ) + 1;
          glp_set_row_bnds(lp, row, GLP_LO, lb, 0.0);
        }
      }
    }
  }

  // (3) Reserve rows (>=) if used
  if (useReserve){
    for (int r=0; r<R; ++r){
      for (int s=0; s<S; ++s){
        for (int h=0; h<H; ++h){
          int tau = s*H + h;
          // RHS = (1+RM)*Demand[r,tau] - Sum_e CF*hrs*Existing[r,e]
          double rhs = (1.0 + ReserveMargin) * Demand[idx_dem(r,s,h,S,H)];
          double hrs = Hours[idx_hours(s,h,H)];
          for (int e=0; e<E; ++e){
            rhs -= CF[idx_cf(r,e,s,h,E,S,H)] * hrs * ExistingCap[idx_xcap(r,e,E)];
          }
          int row = ROW_RSV0 + r*T + tau + 1;
          glp_set_row_bnds(lp, row, GLP_LO, rhs, 0.0);
        }
      }
    }
  }

  // (4) CES row (>=) if used
  if (useCES){
    glp_set_row_bnds(lp, ROW_CES + 1, GLP_LO, CES_rhs, 0.0);
  }

  // Columns (+objective, bounds)
  glp_add_cols(lp, g_N);
  for (int r=0; r<R; ++r){
    int base = r * g_colsPerReg;

    // Build vars
    for (int e=0; e<E; ++e){
      int col = base + col_build_e(e);
      glp_set_obj_coef(lp, col+1, CAPEX[e] + FOM[e]);
      glp_set_col_bnds(lp, col+1, GLP_LO, 0.0, 0.0);
    }
    // Gen vars
    for (int e=0; e<E; ++e){
      double c_e = Fuel[e] + VOM[e];
      for (int tau=0; tau<T; ++tau){
        int col = base + col_gen_e_tau(E, e, tau, T);
        glp_set_obj_coef(lp, col+1, c_e);
        glp_set_col_bnds(lp, col+1, GLP_LO, 0.0, 0.0);
      }
    }
  }

  // Sparse triplets
  std::vector<int> ia(1), ja(1); std::vector<double> ar(1);

  for (int r=0; r<R; ++r){
    int base = r * g_colsPerReg;

    for (int s=0; s<S; ++s){
      for (int h=0; h<H; ++h){
        int tau = s*H + h;

        // Demand row: Sum_e Gen[r,e,tau] = Demand[r,tau]
        int rowD = ROW_DEM0 + r*T + tau + 1;
        for (int e=0; e<E; ++e){
          int colG = base + col_gen_e_tau(E, e, tau, T);
          ia.push_back(rowD); ja.push_back(colG+1); ar.push_back(1.0);
        }

        // Capacity rows: for each e
        double hrs = Hours[idx_hours(s,h,H)];
        for (int e=0; e<E; ++e){
          double cf = CF[idx_cf(r,e,s,h,E,S,H)];
          int rowC = ROW_CAP0 + ((r*E + e)*T + tau) + 1;
          int colG = base + col_gen_e_tau(E, e, tau, T);
          int colB = base + col_build_e(e);
          ia.push_back(rowC); ja.push_back(colG+1); ar.push_back(-1.0);
          ia.push_back(rowC); ja.push_back(colB+1); ar.push_back(cf * hrs);
        }

        // Reserve rows: Sum_e (cf*hrs*Build[r,e]) >= RHS  (if used)
        if (useReserve){
          int rowR = ROW_RSV0 + r*T + tau + 1;
          for (int e=0; e<E; ++e){
            double cf = CF[idx_cf(r,e,s,h,E,S,H)];
            int colB = base + col_build_e(e);
            ia.push_back(rowR); ja.push_back(colB+1); ar.push_back(cf * hrs);
          }
        }

        // CES row: Sum_{r,e,tau} q[e]*Gen[r,e,tau] >= CES_rhs
        if (useCES){
          int rowC = ROW_CES + 1;
          for (int e=0; e<E; ++e){
            double q = CES_q[e];
            if (q != 0.0){
              int colG = base + col_gen_e_tau(E, e, tau, T);
              ia.push_back(rowC); ja.push_back(colG+1); ar.push_back(q);
            }
          }
        }
      }
    }
  }

  glp_load_matrix(lp, (int)ia.size()-1, ia.data(), ja.data(), ar.data());

  g_lastCode = glp_simplex(lp, nullptr);
  if (g_lastCode == 0){
    g_status = glp_get_status(lp);
    g_obj    = glp_get_obj_val(lp);

    // Extract solution
    for (int j=0; j<g_N; ++j) g_solution[j] = glp_get_col_prim(lp, j+1);
    g_solved = (g_status == GLP_OPT || g_status == GLP_FEAS);

    // Per-region costs
    std::fill(g_cost_cap.begin(),    g_cost_cap.end(),    0.0);
    std::fill(g_cost_energy.begin(), g_cost_energy.end(), 0.0);

    for (int r=0; r<R; ++r){
      int base = r * g_colsPerReg;

      // Capacity part
      for (int e=0; e<E; ++e){
        int colB = base + col_build_e(e);
        double build = g_solution[colB];
        g_cost_cap[r] += (CAPEX[e] + FOM[e]) * build;
      }
      // Energy part
      for (int e=0; e<E; ++e){
        double varc = Fuel[e] + VOM[e];
        for (int tau=0; tau<T; ++tau){
          int colG = base + col_gen_e_tau(E, e, tau, T);
          g_cost_energy[r] += varc * g_solution[colG];
        }
      }
    }
  }

  glp_delete_prob(lp);
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

    case 1: *sym=(char*)"LP_CAP_ADD";   *arglist=(char*)" region_pos , tech_pos ";                   *num_args=2; *num_vector=0; *func_index=F_CAP_ADD; return 1;
    case 2: *sym=(char*)"LP_GEN";       *arglist=(char*)" region_pos , tech_pos , ts_pos , hr_pos ";  *num_args=4; *num_vector=0; *func_index=F_GEN;     return 1;
    case 3: *sym=(char*)"LP_COST_CAP";  *arglist=(char*)" region_pos ";                               *num_args=1; *num_vector=0; *func_index=F_COST_CAP;return 1;
    case 4: *sym=(char*)"LP_COST_ENERGY";*arglist=(char*)" region_pos ";                              *num_args=1; *num_vector=0; *func_index=F_COST_EN; return 1;

    case 5: *sym=(char*)"LP_OBJ";       *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_OBJ;    return 1;
    case 6: *sym=(char*)"LP_STATUS";    *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_STATUS; return 1;
    case 7: *sym=(char*)"LP_CODE";      *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_CODE;   return 1;

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
