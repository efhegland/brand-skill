# M1 Compliance — FINRA 2210

**Core requirement:** All retail communications must be fair, balanced, and not misleading.

---

## What you must NEVER do

**1. Make guarantees or promises about performance**
- ❌ "You will save money"
- ❌ "Grow your wealth faster"
- ✅ "Lower fees can improve long-term returns"
- ✅ "May help reduce costs over time"

**2. Tell clients what they SHOULD do**
- ❌ "You should roll over your 401(k)"
- ❌ "You need to use margin for taxes"
- ✅ "A rollover may make sense if..."
- ✅ "Some clients use margin for tax bills"

**3. Create urgency through fear or pressure**
- ❌ "Don't miss out"
- ❌ "Last chance"
- ❌ "You're losing money every day"
- ✅ "Your promotional rate expires [date]"
- ✅ "This rate is available until [date]"

**4. Make unsubstantiated comparisons**
- ❌ "Better than other brokerages"
- ❌ "The best platform for investors"
- ✅ "M1 manages nearly 10 billion dollars in assets" (factual)
- ✅ "Investopedia named M1 the best brokerage for sophisticated investors" (third-party attribution)

**5. Omit material risks**
- ❌ Promote margin without risk disclosure
- ❌ Discuss investment returns without mentioning loss potential
- ✅ Always include: "All investing involves risk, including the risk of losing the money you invest"
- ✅ For margin: "Margin borrowing involves significant risk and isn't suitable for everyone"

---

## What you must ALWAYS do

**1. Use balanced language**
- Show both benefits AND limitations when relevant
- "Lower-cost ETFs may reduce fees, but can differ in strategy and holdings"
- Acknowledge when a product "may not be right for everyone"

**2. Be specific and factual**
- Use precise terms: "expense ratio" not "hidden fees"
- Include actual numbers when making claims: "[4.00%] APY" not "high yield"
- Attribute claims to sources when possible

**3. Include proper risk disclosures**

*For margin/borrowing content:*
- Must mention: increased risk, potential for losses exceeding investment, maintenance calls
- Full disclosure language must appear in every margin communication

*For investment content:*
- Must include: "All investing involves risk, including the risk of losing the money you invest"
- "Past performance does not guarantee future performance"

**4. Avoid promissory language**
- "Can help" not "will help"
- "May reduce" not "reduces"
- "Typically" or "generally" when describing outcomes

**5. Make it clear when advice is needed**
- "You may want to consult a financial advisor"
- "Consider speaking with a tax professional"
- Never position M1 communications as personalized advice

---

## Reference disclosures

Coded disclosures to drop in verbatim when the trigger condition applies. Don't paraphrase — use the exact language below. Full library last synced from Compliance 2026-07-09.

**These disclosures apply to web, email, and blog copy.** They're not required on shorter-form content — product copy, push notifications (PNs), and in-app messages (IAMs) — where there's no room for them to render properly.

Jump to: [General](#general) · [M1 Advisor](#m1-advisor) · [M1 Invest](#m1-invest) · [M1 Margin](#m1-margin) · [M1 Cash Account](#m1-cash-account) · [M1 Savings](#m1-savings) · [M1 Personal Loans](#m1-personal-loans) · [M1 Crypto](#m1-crypto) · [Owner's Rewards Card](#owners-rewards-card)

### General

**General_TE** — present on all emails whose responses aren't monitored (transactional and marketing campaigns):
> This is an automated message, and we will not receive your response if you reply.

**General_DS** — present on all emails/communications where we commingle products or services of M1:
> M1 is a technology company offering a range of financial products and services. "M1" refers to M1 Holdings Inc., and its wholly-owned, separate affiliates M1 Finance LLC, M1 Spend LLC, and M1 Digital LLC.

**General_TM** — anytime we refer to or display trademarks of other companies:
> All product and company names are trademarks™ or registered® trademarks of their respective holders. Use of them does not imply any affiliation with or endorsement by them.

**General_COP** — all marketing emails produced by M1 Holdings:
> © Copyright %%xtyear%% M1 Holdings Inc.

*(`%%xtyear%%` is an ESP merge tag — keep it literal, don't hardcode a year.)*

**Platform_FEE** — present anytime referring to the M1 Platform fee (asterisk in marketing copy, info tooltip in-app/on web):
> † A $3 monthly platform fee will apply to clients with less than $10,000 in M1 assets or without an active Personal Loan. The fee will be waived if your opened M1 Invest or Earn accounts settled aggregate balance equals or exceeds $10,000 for at least one day during the 30-day billing period. The monthly platform fee will be waived for all clients with an active M1 Personal Loan, regardless of their M1 Invest or Earn balances.

### M1 Advisor

**Advisor_GEN** — use any time copy refers to M1 Advisor:
> Advisory products and services are offered by M1 Advisory Services, LLC, an investment adviser registered with the U.S. Securities and Exchange Commission (SEC). Advisory services are distinct from the brokerage products and services offered by M1 Finance LLC. Clients receiving advisory services must also maintain brokerage accounts with M1 Finance LLC, which is a separate legal entity and a Member FINRA/SIPC. For important information about M1 Advisory Services, LLC, including fees, services, and conflicts of interest, please review our Form ADV Part 2A and Form CRS.

**Advisor_EX** — use when showing an example of M1 Advisor chat for marketing purposes:
> This video is an illustrative example using hypothetical figures. It is not based on an actual client, and no actual client interaction is depicted. M1 Advisor is an advisory service provided by M1 Advisory Services, LLC, an SEC-registered investment adviser.

*Advisor_GEN and Advisor_EX also live in the `m1-advisor` skill's `compliance.md`, which is the primary source for Advisor-specific guardrails — kept here too so they're covered even if `m1-advisor` isn't loaded.*

### M1 Invest

**Invest_NR1** — include on any email that names specific stocks, ETFs, or other investment vehicles; makes clear it's informational only, not a recommendation:
> M1 Finance, in its capacity as a brokerage firm, does not provide personalized investment, financial, legal, or tax advice in connection with this communication. Any securities or investment products mentioned are provided for informational purposes only and should not be considered a recommendation or a solicitation to buy or sell any security. You should consult your personal investment, legal, and tax advisors before making any investment decisions. Past performance does not guarantee future results.

**Invest_IR1** — present on all emails where we refer to M1 Invest or investing:
> All investing involves risk, including the risk of losing the money you invest. Brokerage products and services are offered by M1 Finance LLC, Member FINRA / SIPC, and a wholly owned subsidiary of M1 Holdings, Inc.

**Invest_BS** — same disclosure as Invest_IR1, without clickable links; use in plain-text contexts where the Invest_IR1 links (FINRA / SIPC) can't render:
> All investing involves risk, including the risk of losing the money you invest. Brokerage products and services are offered by M1 Finance LLC, Member FINRA / SIPC, and a wholly owned subsidiary of M1 Holdings, Inc.

**Invest_CF^** — anytime we mention commission-free, zero commissions, etc.; link the claim to this disclosure:
> ^M1 Finance, LLC does not charge commission, trading, or management fees for self-directed brokerage accounts. You may still be charged other fees such as M1's platform fee, regulatory fees, account closure fees, or ADR fees. For a complete list of fees M1 may charge visit M1 Fee Schedule.

**Invest_HE** — anytime M1 uses images or screenshots for educational purposes that show securities or strategies:
> All examples above are hypothetical, do not reflect any specific investments, are for informational purposes only, and should not be considered an offer to buy or sell any products. M1 does not provide any financial advice.

**Invest_FS1** — use when referring to fractional shares, to notify customers these shares aren't transferable:
> ªIf you choose to transfer your account to another broker-dealer, only the full shares are guaranteed to transfer. Fractional shares may need to be liquidated and transferred as cash.

**Invest_TW1** — present anytime mentioning AM or PM trade windows or the second trading window; add the ¤ symbol at the end of that AM/PM sentence to link to this disclosure:
> ¤ Participate in both trade windows when you have $25,000 or more equity to comply with pattern-day trading regulations.

**Invest_MS** — anytime we mention the move-a-slice feature:
> Utilizing the Moving Slices feature may cause trades to take place in other accounts you use the same Pie in. Please visit our article on Moving Slices when Pies are used in more than one account for further information.

**N/A (hypothetical performance example)** — anytime showcasing performance on the security graph; bolded values are configurable per situation:
> Hypothetical example for illustrative purpose only. Calculations assuming the following constraints: (a) an initial investment of $1000, (b) an annual rate of return of 10%, (c) no taxes, fees, inflation, or withdrawals. The assumed rate of return is not guaranteed as investing involves risk of loss. Source: [ ]

### M1 Margin

**Margin_MR** — use whenever referring to M1 Margin Loan in general. In-text links: Help Center, M1 APEX margin account risk disclosure, M1 margin account risk disclosure, FINRA/SIPC:
> Brokerage accounts on the M1 platform are either fully disclosed to APEX Clearing or cleared through M1 Finance LLC.
>
> All investing involves risk, including the risk of losing the money you invest. Using margin can add to these risks. Margin rates may vary. Users utilizing APEX cleared margin accounts should review the APEX margin account risk disclosure before borrowing. Users utilizing M1 cleared margin accounts should review the M1 margin account risk disclosure before borrowing. M1 Margin Loans are available on margin accounts with at least $2,000 invested per account. Not available for Retirement or Custodial accounts. Past performance does not guarantee future performance.

### M1 Cash Account

**CashAcct_GEN** — anytime we refer to M1 Cash Accounts, for entity clarity:
> M1 High-Yield Cash Account(s) is an investment product offered by M1 Finance, LLC, an SEC registered broker-dealer, Member FINRA / SIPC. M1 is not a bank and M1 High-Yield Cash Accounts are not a checking or savings account. The purpose of this account is to invest in securities, and an open M1 Investment account is required to participate in the M1 High-Yield Cash Account. All investing involves risk, including the risk of losing the money you invest.

**CashAcct_APY** — anytime mentioning APY for Cash Accounts:
> ¹ Stated APY (annual percentage yield) with the M1 High-Yield Cash Account is accrued on account balance. APY is solely determined by M1 Finance LLC and its partner banks, and will include administrative and account fees that may reduce earnings. Rates are subject to change without notice. M1 High-Yield Cash Account is a separate offering from, and not linked to, the M1 High Yield Savings Accounts offered by M1 Spend LLC's banking partner. M1 is not a bank.

**Cash_Savings_APY** — anytime mentioning APY for "high-yield accounts" and mixing Cash Account with Savings Account APY:
> ¹ M1 is not a bank. M1 High-Yield Cash Account and M1 High Yield Savings Accounts are separate offerings, with varying insurance coverage, and are not linked accounts. Obtaining stated APY (annual percentage yield) or opening the M1 High-Yield Savings account or M1 High-Yield Cash Account does not require a minimum account balance. Stated APY is accrued on account balance and solely determined by M1 Finance, LLC and its partner banks, and will include administrative fees. Account fees for each account type may reduce earning. Rates are subject to change without prior notice. M1 High-Yield Cash Accounts are offered by M1 Finance, LLC, an SEC registered broker-dealer, Member FINRA / SIPC. M1 Spend is a wholly-owned operating subsidiary of M1 Holdings Inc. M1 High-Yield Savings Accounts are furnished by B2 Bank, NA, Member FDIC.

**CashAcct_FDIC** — anytime mentioning insurance for M1 High-Yield Cash Accounts (link: [participating bank list](https://m1.com/legal/agreements/cash-deposit-network/participating-bank-list/)):
> ² The cash balance in your Cash Account is eligible for FDIC Insurance once it is swept to our partner banks and out of your brokerage account. Until the cash balance is swept to partner banks, the funds are held in a brokerage account and protected by SIPC insurance. Once funds are swept to a partner bank, they are no longer held in your brokerage account and are not protected by SIPC insurance. FDIC insurance is not provided until the funds participating in the sweep program leave your brokerage account and into the sweep program. FDIC insurance is applied at the customer profile level. Customers are responsible for monitoring their total assets at each of the sweep program banks. A complete list of participating program banks can be found here.

### M1 Savings

**Savings_GEN** — anytime we refer to M1 Savings Accounts:
> M1 is not a bank. M1 Spend is a wholly-owned operating subsidiary of M1 Holdings Inc. M1 High-Yield Savings Accounts are furnished by B2 Bank, NA, Member FDIC.

**Savings_APY** — anytime mentioning APY for savings accounts:
> ¹ Obtaining stated APY (annual percentage yield) with the M1 High-Yield Savings Account does not require a minimum account balance. Stated APY is accrued on account balance. APY is solely determined by M1 Spend LLC and its partner banks, and will include account fees that will reduce earnings. Rates are subject to change without notice. M1 High-Yield Savings Account is a separate offering from, and not linked to, the M1 High-Yield Cash Account offered by M1 Finance, LLC. M1 is not a bank.

*Any mention of "paid M1 Plus subscription" alongside this disclosure should link to m1.com/plus/membership.*

**Savings_NA** — anytime referencing how many times savings APY beats the national average (link: [FDIC national rates](https://www.fdic.gov/resources/bankers/national-rates/index.html)):
> ² National average is 0.38% APY as of July 2026. Obtained from the FDIC.

**Savings_5MI** — anytime mentioning insurance for M1 High-Yield Savings Accounts:
> ³ M1 High-Yield Savings Accounts are furnished by B2 Bank, NA, Member FDIC ("B2"). M1 is not a bank. B2 Bank is a member FDIC institution and does not itself provide more than $250,000 of FDIC insurance per legal category of account ownership as described in FDIC regulations. Additional FDIC insurance coverage is provided through B2's Insured Deposit Network Program involving other FDIC insured depository institutions. Deposits may be insured up to $5,000,000 through B2's Insured Deposit Network Program. Full terms of the Program can be found at m1.com/legal/agreements/hysa_agreement and a complete list of participating banks in the program can be found at m1.com/legal/agreements/depositnetwork.

### M1 Personal Loans

**Personal_Loans_B2B** — anytime we refer to M1 Personal Loans:
> M1 is not a bank. M1 Personal Loans are furnished by B2 Bank NA, Member FDIC and Equal Opportunity Lender, and serviced by M1 Spend LLC, a wholly-owned operating subsidiary of M1 Holdings, Inc.

**Personal_Loans_APR** — anytime we refer to APR (e.g., "Loans as low as 7.49% APR*"):
> *Rates are not guaranteed and are subject to change. Not all applicants qualify for the lowest available rate and rates are subject to credit history, income, term of loan, and other factors.

**Personal_Loans_CSP** — must be included anytime we say "without impacting your credit score**":
> **To see which personal loan rates and terms you qualify for, M1 conducts a soft credit check that will not affect your credit score. However, if you choose to proceed and continue your application, M1 will request a hard credit check from one or more consumer reporting agencies, which may affect your credit score.

### M1 Crypto

**Crypto_GM** — anytime we market M1 Crypto to prospects or customers (link: [Crypto Disclosures](https://m1.com/crypto-disclosures)):
> Investing in cryptocurrency comes with significant risk and may not be suitable for everyone. Based on your specific situation and financial condition, carefully consider whether investing in cryptocurrencies is suitable for you. For relevant disclosures and risks, visit Crypto Disclosures.
>
> Crypto services, execution, and custody are provided by Bakkt Crypto Solutions LLC (NMLS ID 1828849) through a software licensing agreement with M1 Digital LLC. Bakkt Crypto Solutions LLC and M1 Digital LLC are not registered broker-dealers or FINRA members and your crypto holdings are not securities and are not FDIC or SIPC insured.
>
> M1 Digital LLC is a wholly separate affiliate of M1 Finance LLC, and neither are involved with the execution or custody of cryptocurrencies.

**Crypto_ODT** — anytime referencing on-demand trading for crypto:
> ° On-demand trading for M1 Plus members only, limited to 10 on-demand trades per calendar month. 24/7 availability is subject to scheduled maintenance.

### Owner's Rewards Card

**Credit_GEN** — anytime we mention the Owner's Rewards Card (links: Cardholder Agreement, Rewards Terms):
> Credit Card not available for US Territory Residents. The Owner's Rewards Card by M1 is Powered by Deserve and issued by Celtic Bank. Review Cardholder Agreement and Rewards Terms for important information about the Owner's Rewards Card by M1.

**Credit_CB1** — anytime credit card cash back is mentioned; put * next to the % so it can be linked to the disclosure:
> *2.5%–10% Owner's Rewards cash back is earned on qualifying purchases based on M1's rewards tiers that can be found here. All Standard Reward purchases receive 1.5% cash back. Owners Rewards and Standard Rewards are subject to a maximum of $200 cash back in aggregate per calendar month. Exclusions may apply. See Rewards Terms for additional information and exclusions.

**Credit_OWB** — anytime we mention, show, or refer to any specific stock that's part of Owner's Rewards (e.g., showing an AAPL image); no asterisk needed:
> For informational purposes only and not a trade recommendation. All product and company names are trademarks or registered trademarks of their respective holders. Use of them does not imply any affiliation with or endorsement by them.

---

## FINRA red flags to watch for

**Performance claims:**
- Any statement that implies guaranteed returns
- Comparisons to benchmarks without context
- Cherry-picked time periods

**Omission of material facts:**
- Promoting a benefit without mentioning associated costs
- Discussing returns without risk
- Highlighting a feature without explaining limitations

**Exaggerated claims:**
- Superlatives without substantiation ("best," "only," "fastest")
- Absolute statements
