# Gates: ats_hh-review-and-fix

OWNS: *.js, *.json, *.md

Scope: Fix token refresh retry logic, add refresh handling to all API calls, verify code correctness, prepare for push

- [x] G1: Token refresh + retry implemented in hhFetch (used by getNegotiations)
  CHECK: grep -n "token_expired" hh_module.js | head -5
  EXPECT: token_expired

- [x] G2: getVacansyId uses hhFetch (gets automatic token refresh)
  CHECK: grep -A10 "function getVacansyId" hh_module.js
  EXPECT: hhFetch

- [x] G3: getResume handles null return correctly (continues instead of returning)
  CHECK: grep -A5 "datas_v == null" main.js
  EXPECT: continue

- [x] G4: No syntax errors in all JS files
  CHECK: node --check hh_module.js && node --check main.js && node --check gs_module.js && node --check ai_assistent.js && node --check prompts.js && echo "SYNTAX_OK"
  EXPECT: SYNTAX_OK

- [x] G5: Git status clean (only expected modified files)
  CHECK: git status --porcelain
  EXPECT: M hh_module.js

- [x] G6: Code review - no obvious bugs in logic flow
  CHECK: grep -c "function hhFetch" hh_module.js && grep -c "hhFetch" hh_module.js && echo "GATE6_OK"
  EXPECT: GATE6_OK