# 💡 Examples

`examples/` holds runnable examples of the supported API, normally as VBA
modules in `examples/modules/`. Layout rules are in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

Examples:

- call only the public facade listed in
  [`docs/PUBLIC_API.txt`](../docs/PUBLIC_API.txt), never `src/core` directly;
- use synthetic data;
- are not tests: assertions belong in `tests/`; and
- are never part of the production import set.

No examples exist yet. The public worksheet functions are exercised on the
template's FormulaTest sheet.
