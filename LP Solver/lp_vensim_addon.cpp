// lp_vensim_addon_dyn.cpp
// Build (MSVC, x64, static GLPK, static CRT):
//   vcpkg install glpk:x64-windows-static
//   cl /nologo /LD /O2 /EHsc /MT "lp_vensim_addon_dyn.cpp" ^
//     /I "%VCPKG_ROOT%\installed\x64-windows-static\include" ^
//     /link /NOLOGO ^
//     /LIBPATH:"%VCPKG_ROOT%\installed\x64-windows-static\lib" glpk.lib ^
//     /OUT:"lp_vensim_addon.dll"
//
// Notes:
// - No glpk.dll required at runtime (static linking).
// - GPL: static link to GLPK makes your DLL GPL-covered; include GLPK license & offer source.

#include <windows.h>
#include <vector>
#include <mutex>
#include <type_traits>

#include <glpk.h>  // static link via glpk.lib

// ===== Vensim external function ABI (minimal) =====
#if defined(_MSC_VER)
  #define VEFCC __stdcall
#else
  #define VEFCC
#endif
typedef double COMPREAL;

// VECTOR_ARG and DIM_INFO per Vensim docs
typedef struct {
  COMPREAL*           vals;
  const COMPREAL*     firstval;
  const void*         dim_info;  // not needed here for contiguous copy
  const char*         varname;
} VECTOR_ARG;

typedef union {           // argument union
  COMPREAL      val;      // numeric
  VECTOR_ARG*   vec;      // vector
  void*         tab;      // lookup (unused)
  char*         literal;  // literal (unused)
  void*         constmat; // const matrix (unused)
  void*         datamat;  // data matrix (unused)
} VV;

// ===== LP state (cached) =====
static std::mutex g_mutex;
static std::vector<double> g_solution;
static int    g_status = -1;
static double g_obj    = 0.0;
static bool   g_solved = false;

static inline void reset_cache(int n = 0){
  g_solution.assign(n>0?n:(int)g_solution.size(), 0.0);
  g_status = -1; g_obj = 0.0; g_solved = false;
}

static double solve_lp_min(
  int n, int m,
  const double* c,
  const double* lb,
  const double* ub,
  const double* A,     // row-major MxN
  const int*    sense, // -1 <=, 0 =, +1 >=
  const double* b
){
  reset_cache(n);

  glp_prob* lp = glp_create_prob();
  glp_set_prob_name(lp, "vensim_lp");
  glp_set_obj_dir(lp, GLP_MIN);

  glp_add_rows(lp, m);
  for (int i=0;i<m;++i){
    int sc = sense? sense[i] : 0;
    if (sc<0)       glp_set_row_bnds(lp,i+1,GLP_UP,0.0,b[i]);
    else if (sc==0) glp_set_row_bnds(lp,i+1,GLP_FX,b[i],b[i]);
    else            glp_set_row_bnds(lp,i+1,GLP_LO,b[i],0.0);
  }

  glp_add_cols(lp, n);
  const double INF = 1e30;
  for (int j=0;j<n;++j){
    glp_set_obj_coef(lp, j+1, c[j]);
    double lo=lb[j], hi=ub[j];
    if (lo<=-INF && hi>=INF) glp_set_col_bnds(lp,j+1,GLP_FR,0.0,0.0);
    else if (hi>=INF)        glp_set_col_bnds(lp,j+1,GLP_LO,lo,0.0);
    else if (lo<=-INF)       glp_set_col_bnds(lp,j+1,GLP_UP,0.0,hi);
    else if (lo==hi)         glp_set_col_bnds(lp,j+1,GLP_FX,lo,hi);
    else                     glp_set_col_bnds(lp,j+1,GLP_DB,lo,hi);
  }

  // sparse triplets
  std::vector<int> ia(1), ja(1);
  std::vector<double> ar(1);
  ia.reserve(1 + (size_t)m*(size_t)n);
  ja.reserve(1 + (size_t)m*(size_t)n);
  ar.reserve(1 + (size_t)m*(size_t)n);
  for (int i=0;i<m;++i){
    for (int j=0;j<n;++j){
      double v = A[i*n + j];
      if (v!=0.0){ ia.push_back(i+1); ja.push_back(j+1); ar.push_back(v); }
    }
  }
  glp_load_matrix(lp, (int)ia.size()-1, ia.data(), ja.data(), ar.data());

  // quiet simplex with defaults
  glp_term_out(GLP_OFF);
  glp_smcp parm; glp_init_smcp(&parm);
  parm.msg_lev = GLP_MSG_OFF;
  int ret = glp_simplex(lp, &parm);

  g_status = ret; // keep original behavior: 0=success, nonzero=solver error
  if (ret==0){
    int stat = glp_get_status(lp);
    g_obj = glp_get_obj_val(lp);
    g_solution.assign(n,0.0);
    for (int j=0;j<n;++j) g_solution[j] = glp_get_col_prim(lp,j+1);
    g_solved = (stat==GLP_OPT || stat==GLP_FEAS);
  }
  glp_delete_prob(lp);
  return g_solved? g_obj : 1e308;
}

// ===== Vensim-required entry points =====

// IMPORTANT: Must match your Vensim external version.
static const int EXTERN_VCODE = 62051;

// 1) version_info — Vensim probes this first
extern "C" __declspec(dllexport) int VEFCC version_info(){
  return EXTERN_VCODE;
}

// function IDs we’ll advertise via user_definition
enum { F_LP_SOLVE=1001, F_LP_X=1002, F_LP_OBJ=1003, F_LP_STATUS=1004 };

// 2) user_definition — advertise functions & their signatures
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

  *dim_act = 0; *modify = 0; *num_loops = 0; *num_literal = 0; *num_lookup = 0;

  switch (setup_index){
    case 0: // LP_SOLVE(c,lb,ub,A_flat,sense,b,N,M)
      *sym        = (char*)"LP_SOLVE";
      *arglist    = (char*)" {c} , {lb} , {ub} , {A_flat} , {sense} , {b} , N , M ";
      *num_args   = 8;  // total typed arguments
      *num_vector = 6;  // first six are vectors
      *func_index = F_LP_SOLVE;
      return 1;

    case 1: // LP_X(idx)
      *sym = (char*)"LP_X";
      *arglist = (char*)" {idx} ";
      *num_vector = 0;
      *num_args   = 1;
      *func_index = F_LP_X;
      return 1;

    case 2: // LP_OBJ()
      *sym = (char*)"LP_OBJ";
      *arglist = (char*)" ";
      *num_vector = 0;
      *num_args   = 0;
      *func_index = F_LP_OBJ;
      return 1;

    case 3: // LP_STATUS()
      *sym = (char*)"LP_STATUS";
      *arglist = (char*)" ";
      *num_vector = 0;
      *num_args   = 0;
      *func_index = F_LP_STATUS;
      return 1;

    default:
      return 0; // done
  }
}

// 3) vensim_external — Vensim calls this to execute functions
extern "C" __declspec(dllexport) int VEFCC vensim_external(VV* val, int nval, int funcid){
  std::lock_guard<std::mutex> lk(g_mutex);

  switch (funcid){
    case F_LP_SOLVE: {
      // [0]=c vec, [1]=lb vec, [2]=ub vec, [3]=A_flat vec, [4]=sense vec, [5]=b vec, [6]=N scalar, [7]=M scalar
      if (nval < 8) { val[0].val = 1e308; return 0; }

      auto v_c     = val[0].vec->vals;
      auto v_lb    = val[1].vec->vals;
      auto v_ub    = val[2].vec->vals;
      auto v_A     = val[3].vec->vals;
      auto v_sense = val[4].vec->vals;
      auto v_b     = val[5].vec->vals;

      int N = (int)(val[6].val + 0.5);
      int M = (int)(val[7].val + 0.5);
      if (N<=0 || M<=0){ val[0].val = 1e308; return 0; }

      std::vector<int> sense(M,0);
      for (int i=0;i<M;++i){
        double s = v_sense[i];
        sense[i] = (s < -0.5 ? -1 : (s > 0.5 ? +1 : 0));
      }

      double obj = solve_lp_min(N, M, v_c, v_lb, v_ub, v_A, sense.data(), v_b);
      val[0].val = obj; // return objective
      return 1;
    }

    case F_LP_X: {
      if (nval < 1) { return 0; }
      int idx = (int)(val[0].val + 0.5); // grab arg before overwriting
      if (!g_solved || idx<=0 || idx>(int)g_solution.size()){
        val[0].val = 1e308;
      } else {
        val[0].val = g_solution[idx-1];
      }
      return 1;
    }

    case F_LP_OBJ: {
      val[0].val = g_obj;
      return 1;
    }

    case F_LP_STATUS: {
      val[0].val = (double)g_status; // 0 if glp_simplex() returned success
      return 1;
    }

    default:
      return 0;
  }
}
