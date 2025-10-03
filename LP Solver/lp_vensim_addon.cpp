// lp_vensim_addon.cpp
// Single-solver Vensim addon using GLPK (static link).
// Exports:
//   LP_SOLVE_IMPLICIT_RCF_QBV_XCAP(Demand,Capacity_factor,Hours,
//                                  CAPEX,FOM,Fuel,VOM,
//                                  ExistingCap,CES_qualifying,CES_rhs,
//                                  R,S,H)
//   LP_CAP_GAS(region_pos),  LP_CAP_WIND(region_pos)
//   LP_GEN_GAS(region_pos, ts_pos, hr_pos)
//   LP_GEN_WIND(region_pos, ts_pos, hr_pos)
//   LP_OBJ(), LP_STATUS(), LP_CODE()
//
// CES row is now CONDITIONAL: omitted entirely if (CES_rhs <= 0) or (all tech weights = 0).
//
// Build:
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
  const COMPREAL* firstval;  // base of array (per Vensim docs)
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
static std::vector<double> g_solution;
static int                 g_status   = -1;
static int                 g_lastCode = -999;
static double              g_obj      = 0.0;
static bool                g_solved   = false;

// Shape of last solve
static int g_R=0, g_S=0, g_H=0;
static int g_T=0, g_colsPerReg=0, g_N=0;

static inline void reset_cache(int n = 0){
  g_solution.assign(n>0?n:(int)g_solution.size(), 0.0);
  g_status = -1; g_lastCode = -999; g_obj = 0.0; g_solved = false;
  g_R=g_S=g_H=g_T=g_colsPerReg=g_N=0;
}

// ---------- Index helpers ----------
static inline int idx_cf(int tech, int r, int s, int h, int R, int S, int H){
  // CF dims [TECH,REGION,TS,H]; last subscript (H) varies fastest
  return (((tech*R + r)*S + s)*H + h);
}
static inline int idx_hours(int s, int h, int H){ return s*H + h; }            // Hours [TS,H]
static inline int idx_dem(int r, int s, int h, int S, int H){ return ((r*S)+s)*H + h; } // Demand [R,TS,H]
static inline int idx_xcap(int tech, int r, int R){ return tech*R + r; }       // ExistingCap [TECH,R]

// ---------- Single solver (implicit RCF + CES + ExistingCap) ----------
static double solve_lp_implicit_rcf_q_xcap(
  const double* Demand,   int R, int S, int H,           // [R,TS,H]
  const double* CF,                                      // [TECH>=2, R, TS, H] (0=gas,1=wind)
  const double* Hours,                                   // [TS,H]
  double CAPEX_g, double FOM_g,  double CAPEX_w, double FOM_w,
  double Fuel_g,  double VOM_g,  double Fuel_w,  double VOM_w,
  double q_gas,   double q_wind,
  const double* ExistingCap,                             // [TECH>=2, R]
  double CES_rhs
){
  const int T = S*H;
  const int ColsPerReg = 2 + 2*T; // Build_g, Build_w, Gen_g(T), Gen_w(T)
  const int N = R * ColsPerReg;

  // Decide whether to include a CES row at all
  const bool use_ces = (CES_rhs > 0.0) && ((q_gas != 0.0) || (q_wind != 0.0));

  const int ROW_DEM   = 0;
  const int ROW_CAP_G = ROW_DEM + R*T;
  const int ROW_CAP_W = ROW_CAP_G + R*T;
  const int ROW_CES   = ROW_CAP_W + R*T;               // only valid if use_ces
  const int M         = ROW_CAP_W + R*T + (use_ces ? 1 : 0);

  reset_cache(N);
  g_R=R; g_S=S; g_H=H; g_T=T; g_colsPerReg=ColsPerReg; g_N=N;

  glp_prob* lp = glp_create_prob();
  glp_set_prob_name(lp, "vensim_lp_impl_rcf_q_xcap");
  glp_set_obj_dir(lp, GLP_MIN);

  glp_add_rows(lp, M);

  // Demand balance (=)
  for (int r=0; r<R; ++r){
    for (int s=0; s<S; ++s){
      for (int h=0; h<H; ++h){
        int t = s*H + h;
        int row = ROW_DEM + r*T + t + 1;
        double rhs = Demand[idx_dem(r,s,h,S,H)];
        glp_set_row_bnds(lp, row, GLP_FX, rhs, rhs);
      }
    }
  }

  // Capacity rows (>= LB with existing cap contribution)
  for (int r=0; r<R; ++r){
    double xcap_g = ExistingCap[idx_xcap(0,r,R)];
    double xcap_w = ExistingCap[idx_xcap(1,r,R)];
    for (int s=0; s<S; ++s){
      for (int h=0; h<H; ++h){
        int t = s*H + h;
        double cf_g = CF[idx_cf(0,r,s,h,R,S,H)];
        double cf_w = CF[idx_cf(1,r,s,h,R,S,H)];
        double hrs  = Hours[idx_hours(s,h,H)];
        double lb_g = -cf_g * hrs * xcap_g; // -gen + cf*hrs*Build >= -cf*hrs*Existing
        double lb_w = -cf_w * hrs * xcap_w;
        glp_set_row_bnds(lp, ROW_CAP_G + r*T + t + 1, GLP_LO, lb_g, 0.0);
        glp_set_row_bnds(lp, ROW_CAP_W + r*T + t + 1, GLP_LO, lb_w, 0.0);
      }
    }
  }

  // CES row (>= target) — only if use_ces
  if (use_ces){
    glp_set_row_bnds(lp, ROW_CES + 1, GLP_LO, CES_rhs, 0.0);
  }

  // Columns (objective + bounds)
  glp_add_cols(lp, N);
  for (int r=0; r<R; ++r){
    int base = r * ColsPerReg;

    // Build gas/wind
    glp_set_obj_coef(lp, base+1, CAPEX_g + FOM_g); glp_set_col_bnds(lp, base+1, GLP_LO, 0.0, 0.0);
    glp_set_obj_coef(lp, base+2, CAPEX_w + FOM_w); glp_set_col_bnds(lp, base+2, GLP_LO, 0.0, 0.0);

    // Gen gas/wind per (s,h)
    for (int t=0; t<T; ++t){
      glp_set_obj_coef(lp, base+2 + t + 1,           Fuel_g + VOM_g);
      glp_set_col_bnds(lp, base+2 + t + 1,           GLP_LO, 0.0, 0.0);
      glp_set_obj_coef(lp, base+2 + T + t + 1,       Fuel_w + VOM_w);
      glp_set_col_bnds(lp, base+2 + T + t + 1,       GLP_LO, 0.0, 0.0);
    }
  }

  // Sparse triplets
  std::vector<int> ia(1), ja(1);
  std::vector<double> ar(1);

  for (int r=0; r<R; ++r){
    int base = r * ColsPerReg;
    for (int s=0; s<S; ++s){
      for (int h=0; h<H; ++h){
        int t = s*H + h;
        int rowD = ROW_DEM + r*T + t + 1;
        int col_gg = base + 2 + t + 1;
        int col_gw = base + 2 + T + t + 1;

        // Demand balance: gen_g + gen_w = Demand
        ia.push_back(rowD); ja.push_back(col_gg); ar.push_back(1.0);
        ia.push_back(rowD); ja.push_back(col_gw); ar.push_back(1.0);

        double cf_g = CF[idx_cf(0,r,s,h,R,S,H)];
        double cf_w = CF[idx_cf(1,r,s,h,R,S,H)];
        double hrs  = Hours[idx_hours(s,h,H)];

        // Gas capacity: -gen_g + CF_g*hrs * build_g >= -CF_g*hrs*existing_g
        int rowG = ROW_CAP_G + r*T + t + 1;
        ia.push_back(rowG); ja.push_back(col_gg); ar.push_back(-1.0);
        ia.push_back(rowG); ja.push_back(base + 1); ar.push_back(cf_g * hrs);

        // Wind capacity
        int rowW = ROW_CAP_W + r*T + t + 1;
        ia.push_back(rowW); ja.push_back(col_gw); ar.push_back(-1.0);
        ia.push_back(rowW); ja.push_back(base + 2); ar.push_back(cf_w * hrs);

        // CES weights (q_gas, q_wind) — only if use_ces
        if (use_ces){
          if (q_gas  != 0.0) { ia.push_back(ROW_CES + 1); ja.push_back(col_gg); ar.push_back(q_gas); }
          if (q_wind != 0.0) { ia.push_back(ROW_CES + 1); ja.push_back(col_gw); ar.push_back(q_wind); }
        }
      }
    }
  }

  glp_load_matrix(lp, (int)ia.size()-1, ia.data(), ja.data(), ar.data());

  g_lastCode = glp_simplex(lp, nullptr);
  if (g_lastCode == 0){
    g_status = glp_get_status(lp);
    g_obj    = glp_get_obj_val(lp);
    g_solution.assign(N, 0.0);
    for (int j=0; j<N; ++j) g_solution[j] = glp_get_col_prim(lp, j+1);
    g_solved = (g_status == GLP_OPT || g_status == GLP_FEAS);
  }
  glp_delete_prob(lp);
  return g_solved ? g_obj : 1e308;
}

// ---------- Vensim exports ----------
static const int EXTERN_VCODE = 62051;
extern "C" __declspec(dllexport) int VEFCC version_info(){ return EXTERN_VCODE; }

// Function IDs
enum {
  F_SOLVE_XCAP = 1001,
  F_CAP_GAS,
  F_CAP_WIND,
  F_GEN_GAS,
  F_GEN_WIND,
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
  if (!sym||!arglist||!num_args||!num_vector||!func_index||!dim_act||!modify||!num_loops||!num_literal||!num_lookup)
    return 0;
  *dim_act=0; *modify=0; *num_loops=0; *num_literal=0; *num_lookup=0;

  switch (setup_index){
    case 0:
      *sym        = (char*)"LP_SOLVE_IMPLICIT_RCF_QBV_XCAP";
      *arglist    = (char*)"{Demand},{Capacity_factor},{Hours},{CAPEX},{FOM},{Fuel},{VOM},{ExistingCap},{CES_qualifying},CES_rhs,R,S,H";
      *num_args   = 13;
      *num_vector = 9;
      *func_index = F_SOLVE_XCAP;
      return 1;

    case 1: *sym=(char*)"LP_CAP_GAS";  *arglist=(char*)" region_pos ";                   *num_args=1; *num_vector=0; *func_index=F_CAP_GAS;  return 1;
    case 2: *sym=(char*)"LP_CAP_WIND"; *arglist=(char*)" region_pos ";                   *num_args=1; *num_vector=0; *func_index=F_CAP_WIND; return 1;
    case 3: *sym=(char*)"LP_GEN_GAS";  *arglist=(char*)" region_pos , ts_pos , hr_pos "; *num_args=3; *num_vector=0; *func_index=F_GEN_GAS;  return 1;
    case 4: *sym=(char*)"LP_GEN_WIND"; *arglist=(char*)" region_pos , ts_pos , hr_pos "; *num_args=3; *num_vector=0; *func_index=F_GEN_WIND; return 1;

    case 5: *sym=(char*)"LP_OBJ";      *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_OBJ;    return 1;
    case 6: *sym=(char*)"LP_STATUS";   *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_STATUS; return 1;
    case 7: *sym=(char*)"LP_CODE";     *arglist=(char*)""; *num_args=0; *num_vector=0; *func_index=F_LP_CODE;   return 1;

    default: return 0;
  }
}

extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch (funcid){

    case F_SOLVE_XCAP: {
      if (nval < 13) { val[0].val = 1e308; return 0; }
      const double* Dem    = val[0].vec->firstval;
      const double* CF     = val[1].vec->firstval;
      const double* Hours  = val[2].vec->firstval;
      const double* CAPEXv = val[3].vec->firstval;
      const double* FOMv   = val[4].vec->firstval;
      const double* Fuelv  = val[5].vec->firstval;
      const double* VOMv   = val[6].vec->firstval;
      const double* XCap   = val[7].vec->firstval;
      const double* CESq   = val[8].vec->firstval;
      double CESrhs        = val[9].val;
      int R = (int)(val[10].val + 0.5);
      int S = (int)(val[11].val + 0.5);
      int H = (int)(val[12].val + 0.5);

      double CAPEX_g = CAPEXv[0], CAPEX_w = CAPEXv[1];
      double FOM_g   = FOMv[0],   FOM_w   = FOMv[1];
      double Fuel_g  = Fuelv[0],  Fuel_w  = Fuelv[1];
      double VOM_g   = VOMv[0],   VOM_w   = VOMv[1];
      double q_gas   = CESq[0],   q_wind  = CESq[1];

      val[0].val = solve_lp_implicit_rcf_q_xcap(Dem,R,S,H,CF,Hours,
                                                CAPEX_g,FOM_g,CAPEX_w,FOM_w,
                                                Fuel_g,VOM_g,Fuel_w,VOM_w,
                                                q_gas,q_wind,
                                                XCap,
                                                CESrhs);
      return 1;
    }

    // ----- Scalar getters (1-based positions) -----
    case F_CAP_GAS: {
      if (!g_solved || nval<1) { val[0].val=1e308; return 1; }
      int r = (int)(val[0].val + 0.5); if (r<1 || r>g_R){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      val[0].val = g_solution[ base + 0 ]; // build_g
      return 1;
    }
    case F_CAP_WIND: {
      if (!g_solved || nval<1) { val[0].val=1e308; return 1; }
      int r = (int)(val[0].val + 0.5); if (r<1 || r>g_R){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      val[0].val = g_solution[ base + 1 ]; // build_w
      return 1;
    }
    case F_GEN_GAS: {
      if (!g_solved || nval<3) { val[0].val=1e308; return 1; }
      int r  = (int)(val[0].val + 0.5);
      int ts = (int)(val[1].val + 0.5);
      int hr = (int)(val[2].val + 0.5);
      if (r<1||r>g_R||ts<1||ts>g_S||hr<1||hr>g_H){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      int t = (ts-1)*g_H + (hr-1);
      int idx = base + 2 + t; // zero-based
      val[0].val = g_solution[idx];
      return 1;
    }
    case F_GEN_WIND: {
      if (!g_solved || nval<3) { val[0].val=1e308; return 1; }
      int r  = (int)(val[0].val + 0.5);
      int ts = (int)(val[1].val + 0.5);
      int hr = (int)(val[2].val + 0.5);
      if (r<1||r>g_R||ts<1||ts>g_S||hr<1||hr>g_H){ val[0].val=1e308; return 1; }
      int base = (r-1)*g_colsPerReg;
      int t = (ts-1)*g_H + (hr-1);
      int idx = base + 2 + g_T + t; // zero-based
      val[0].val = g_solution[idx];
      return 1;
    }

    // ----- Diagnostics -----
    case F_LP_OBJ:    { val[0].val = g_obj;       return 1; }
    case F_LP_STATUS: { val[0].val = (double)g_status;   return 1; }
    case F_LP_CODE:   { val[0].val = (double)g_lastCode; return 1; }

    default: return 0;
  }
}
