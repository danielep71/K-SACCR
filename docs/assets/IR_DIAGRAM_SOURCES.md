# IR add-on diagram corrections

The SVG is the source for `SACCR_IR_AddOn_flow.png`. This documentation diagram
does not certify the engine or replace the methodology source-register decisions.

- **Supervisory duration:** Basel CRE52.34 floors SD at ten business days.
  CRR Article 279b(1)(a) also has this floor following the corrigendum to
  Regulation (EU) 2019/876, OJ L 398, 11 November 2021, item (14).
- **Maturity factor:** CRE52.48 and CRR Article 279c(1)(a) use a ten-business-day
  maturity floor and a one-year cap for unmargined trades. `B` denotes business
  days in a year under the applicable convention; dates and SD are in years.
- **Buckets:** CRE52.57 step 3 puts E < 1 year in bucket 1. CRR Article 280a,
  Table 2, puts 0 < E <= 1 year in bucket 1 and 1 < E <= 5 in bucket 2.
- **Inflation:** CRE52.57 FAQ1 and CRR Article 277a(1) separate inflation
  hedging sets by currency from ordinary interest-rate hedging sets.

Primary texts inspected:

1. [Basel CRE52, effective 1 January 2023, published 5 June 2020](https://www.bis.org/committees/bcbs/basel-framework/standard/cre/52/inforce/2023-01-01/published/2020-06-05).
2. [Regulation (EU) 2019/876, official legislation.gov.uk copy](https://www.legislation.gov.uk/eur/2019/876/2023-07-11?view=plain), inserted Articles 277a, 279c and 280a.
3. [Operative corrigendum, OJ L 398, 11 November 2021](https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX%3A32019R0876R%2811%29), item (14), printed pages 36–37. The earlier OJ L 65 corrigendum of 25 February 2021 was [declared null and void on 27 October 2021](https://eur-lex.europa.eu/eli/reg/2019/876/corrigendum/2021-10-27/oj/eng); use the November reissue. The reissued Official Journal text was checked using [this full Journal PDF](https://urlausnir.is/skrar/Stj%C3%B3rnart%C3%AD%C3%B0indi%20ESB/2021/Stj%C3%B3rnart%C3%AD%C3%B0indi_ESB_2021_L398_EN.pdf); EUR-Lex served a robot-verification page during this check.

Regenerate with Inkscape:

```sh
inkscape docs/assets/SACCR_IR_AddOn_flow.svg --export-type=png --export-width=2080 --export-filename=docs/assets/SACCR_IR_AddOn_flow.png
```
