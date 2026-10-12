# ⚖️ Basel vs CRR

> **Guide, not policy:** [Methodology](../methodology/README.md#regime-differences) owns the registered comparison and exact locators.

## Read this as an implementation map

This page summarizes the current repository decisions. It replaces the prototype comparison, including its obsolete claim that the ordinary margined cap uses VM under Basel but excludes it under CRR. It is not a new independent legal opinion.

| Topic | Current project treatment |
| --- | --- |
| Margined exposure cap | Uses NICA on the hypothetical unmargined cap basis in both regimes, following #39. The former BCBS use of VM + NICA is superseded |
| One-way margining | Handled separately from the ordinary cap; follow the support matrix and regime-specific input contract |
| Interest-rate aggregation | Offset formula supported in both regimes; the no-offset choice is allowed only under BCBS |
| Other risks | CRR OT category supported; OT trades rejected under BCBS |
| Climatic commodities | Separate CRR hedging set; the registered 18% factor is no longer an unconfirmed placeholder |
| Credit classification | Rating and CQS mappings remain distinct; user supplies the required classification |
| Option lambda | CRR calculation follows the registered RTS treatment; BCBS follows its separate supported input contract |
| Economic trade mapping | User supplies category/component mapping; the engine checks grouped identity, netting set, and the one-MtM-leg rule |
| Option structures | Supported within the documented decomposition boundary; splitting a linear product into multiple legs of the same class is rejected |
| Digital options | Generic option-leg support is not proof of digital payoff sizing or cap correctness; [#120](https://github.com/danielep71/VBA-SACCR-Toolkit/issues/120) remains open |
| Sold options/protection | Optional zero-EAD and unpaid-premium reductions are not automatically applied |
| Multiple/shared margin agreements | Unsupported declared structures are rejected; they are not silently calculated as ordinary sets |

The registered comparison provides the relevant CRR articles and CRE52 paragraphs. It also explains why the cap is no longer listed as a regime difference.

## Avoid three misleading comparisons

1. A same-number result does not establish that two legal rules are identical.
2. A derived calculation is not a published benchmark.
3. A passing input or formula test is not complete portfolio or regulatory validation.

See [Worked Examples](Worked-Examples.md) for units and evidence classes.
