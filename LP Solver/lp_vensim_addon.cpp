// lp_vensim_addon_dyn.cpp
// Build (MSVC): cl /LD "C:\Users\RobbieOrvis\Downloads\Vensim LP\lp_vensim_addon_dyn.cpp" /O2 /EHsc /MT ^
//   /link /DEF:"C:\Users\RobbieOrvis\Downloads\Vensim LP\lp_vensim_addon.def"
// Place glpk.dll (or libglpk-0.dll) next to the built DLL.

#include <windows.h>
#include <vector>
#include <mutex>
#include <type_traits>

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

// ===== GLPK dynamic binding (no headers needed) =====
struct glp_prob;
static HMODULE hGLPK = nullptr;

static const int GLP_MIN = 1;
static const int GLP_FR  = 1, GLP_LO = 2, GLP_UP = 3, GLP_DB = 4, GLP_FX = 5;
static const int GLP_OPT = 5, GLP_FEAS = 2;

typedef glp_prob* (__cdecl *PFN_glp_create_prob)(void);
typedef void      (__cdecl *PFN_glp_set_prob_name)(glp_prob*, const char*);
typedef void      (__cdecl *PFN_glp_set_obj_dir)(glp_prob*, int);
typedef void      (__cdecl *PFN_glp_add_rows)(glp_prob*, int);
typedef void      (__cdecl *PFN_glp_set_row_bnds)(glp_prob*, int, int, double, double);
typedef void      (__cdecl *PFN_glp_add_cols)(glp_prob*, int);
typedef void      (__cdecl *PFN_glp_set_obj_coef)(glp_prob*, int, double);
typedef void      (__cdecl *PFN_glp_set_col_bnds)(glp_prob*, int, int, double, double);
typedef void      (__cdecl *PFN_glp_load_matrix)(glp_prob*, int, const int[], const int[], const double[]);
typedef int       (__cdecl *PFN_glp_simplex)(glp_prob*, const void*);
typedef int       (__cdecl *PFN_glp_get_status)(glp_prob*);
typedef double    (__cdecl *PFN_glp_get_obj_val)(glp_prob*);
typedef double    (__cdecl *PFN_glp_get_col_prim)(glp_prob*, int);
typedef void      (__cdecl *PFN_glp_delete_prob)(glp_prob*);

static PFN_glp_create_prob   p_glp_create_prob   = nullptr;
static PFN_glp_set_prob_name p_glp_set_prob_name = nullptr;
static PFN_glp_set_obj_dir   p_glp_set_obj_dir   = nullptr;
static PFN_glp_add_rows      p_glp_add_rows      = nullptr;
static PFN_glp_set_row_bnds  p_glp_set_row_bnds  = nullptr;
static PFN_glp_add_cols      p_glp_add_cols      = nullptr;
static PFN_glp_set_obj_coef  p_glp_set_obj_coef  = nullptr;
static PFN_glp_set_col_bnds  p_glp_set_col_bnds  = nullptr;
static PFN_glp_load_matrix   p_glp_load_matrix   = nullptr;
static PFN_glp_simplex       p_glp_simplex       = nullptr;
static PFN_glp_get_status    p_glp_get_status    = nullptr;
static PFN_glp_get_obj_val   p_glp_get_obj_val   = nullptr;
static PFN_glp_get_col_prim  p_glp_get_col_prim  = nullptr;
static PFN_glp_delete_prob   p_glp_delete_prob   = nullptr;

static bool ensure_glpk_loaded() {
  if (hGLPK) return true;
  const char* cands[] = {"glpk.dll","libglpk-0.dll"};
  for (auto n : cands) { hGLPK = LoadLibraryA(n); if (hGLPK) break; }
  if (!hGLPK) return false;
  auto req = [&](auto& fn, const char* name){
    fn = reinterpret_cast<std::remove_reference_t<decltype(fn)>>(GetProcAddress(hGLPK, name));
    return fn != nullptr;
  };
  bool ok = true;
  ok &= req(p_glp_create_prob,   "glp_create_prob");
  ok &= req(p_glp_set_prob_name, "glp_set_prob_name");
  ok &= req(p_glp_set_obj_dir,   "glp_set_obj_dir");
  ok &= req(p_glp_add_rows,      "glp_add_rows");
  ok &= req(p_glp_set_row_bnds,  "glp_set_row_bnds");
  ok &= req(p_glp_add_cols,      "glp_add_cols");
  ok &= req(p_glp_set_obj_coef,  "glp_set_obj_coef");
  ok &= req(p_glp_set_col_bnds,  "glp_set_col_bnds");
  ok &= req(p_glp_load_matrix,   "glp_load_matrix");
  ok &= req(p_glp_simplex,       "glp_simplex");
  ok &= req(p_glp_get_status,    "glp_get_status");
  ok &= req(p_glp_get_obj_val,   "glp_get_obj_val");
  ok &= req(p_glp_get_col_prim,  "glp_get_col_prim");
  ok &= req(p_glp_delete_prob,   "glp_delete_prob");
  if (!ok) { FreeLibrary(hGLPK); hGLPK = nullptr; }
  return ok;
}

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
  if (!ensure_glpk_loaded()) return 1e308;
  reset_cache(n);

  glp_prob* lp = p_glp_create_prob();
  p_glp_set_prob_name(lp, "vensim_lp");
  p_glp_set_obj_dir(lp, GLP_MIN);

  p_glp_add_rows(lp, m);
  for (int i=0;i<m;++i){
    int sc = sense? sense[i] : 0;
    if (sc<0)       p_glp_set_row_bnds(lp,i+1,GLP_UP,0.0,b[i]);
    else if (sc==0) p_glp_set_row_bnds(lp,i+1,GLP_FX,b[i],b[i]);
    else            p_glp_set_row_bnds(lp,i+1,GLP_LO,b[i],0.0);
  }

  p_glp_add_cols(lp, n);
  const double INF = 1e30;
  for (int j=0;j<n;++j){
    p_glp_set_obj_coef(lp, j+1, c[j]);
    double lo=lb[j], hi=ub[j];
    if (lo<=-INF && hi>=INF) p_glp_set_col_bnds(lp,j+1,GLP_FR,0.0,0.0);
    else if (hi>=INF)        p_glp_set_col_bnds(lp,j+1,GLP_LO,lo,0.0);
    else if (lo<=-INF)       p_glp_set_col_bnds(lp,j+1,GLP_UP,0.0,hi);
    else if (lo==hi)         p_glp_set_col_bnds(lp,j+1,GLP_FX,lo,hi);
    else                     p_glp_set_col_bnds(lp,j+1,GLP_DB,lo,hi);
  }

  // sparse triplets
  std::vector<int> ia(1), ja(1);
  std::vector<double> ar(1);
  ia.reserve(1+m*n); ja.reserve(1+m*n); ar.reserve(1+m*n);
  for (int i=0;i<m;++i){
    for (int j=0;j<n;++j){
      double v = A[i*n + j];
      if (v!=0.0){ ia.push_back(i+1); ja.push_back(j+1); ar.push_back(v); }
    }
  }
  p_glp_load_matrix(lp, (int)ia.size()-1, ia.data(), ja.data(), ar.data());

  int ret = p_glp_simplex(lp, nullptr);
  g_status = ret;
  if (ret==0){
    int stat = p_glp_get_status(lp);
    g_obj = p_glp_get_obj_val(lp);
    g_solution.assign(n,0.0);
    for (int j=0;j<n;++j) g_solution[j] = p_glp_get_col_prim(lp,j+1);
    g_solved = (stat==GLP_OPT || stat==GLP_FEAS);
  }
  p_glp_delete_prob(lp);
  return g_solved? g_obj : 1e308;
}

// ===== Vensim-required entry points =====

// IMPORTANT: This value must match your Vensim version.
// Try 5841 first; if Vensim still refuses to load, recompile with 5840.
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
  // Initialize outputs to safe defaults
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
	  *dim_act    = 0; *modify = 0; *num_loops = 0; *num_literal = 0; *num_lookup = 0;
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
      // Argument order in VV:
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

      // convert sense to ints
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
      val[0].val = (double)g_status;
      return 1;
    }

    default:
      return 0;
  }
}
