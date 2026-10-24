# web scraping excel: The Built-In Free Methods vs. a Scraping API, With Real Costs for Each Route

Getting website data into Excel comes down to three routes: Power Query for static HTML tables, VBA for custom parsing, and a scraping API when the page is JavaScript-rendered or protected by anti-bot systems. Most people searching this are stuck at the same fork: the free built-in tools work great on Wikipedia-style tables, then silently return empty results on the e-commerce or listing page they actually care about.

This guide covers how each method works, where each one breaks, and what the paid route costs in practice — including ScraperAPI's current plans ($49/month entry, 7-day trial with 5,000 free credits, a 7-day no-questions-asked refund) and the per-site credit pricing most reviews skip entirely.

## Quick take

If your target page has a real HTML table, stop reading and use Excel's **Data > From Web** — it's free and takes two minutes. If the site loads content with JavaScript, sits behind Cloudflare or similar protection, or you need scheduled bulk pulls into a spreadsheet, a scraping API is the practical option. ScraperAPI's Hobby plan at **$49/month** (or $44.10/month billed annually) covers 100,000 credits with 20 concurrent threads, and you can test whether it works on your specific target site before paying anything.

👉 [Start the 7-day free trial and compare all plans](https://www.scraperapi.com/?fp_ref=coupons)

## Method 1: Power Query — free, built into Excel, and where it fails

Power Query is Excel's native web connector, and it should always be your first attempt because it costs nothing.

1. Open Excel and go to **Data > From Web**
2. Paste the target URL (a Wikipedia GDP list is the classic test case)
3. Excel detects the HTML tables on the page automatically
4. Select the table you want and click **Load**

Two features make this more useful than most tutorials suggest. The query is refreshable — right-click the loaded table to update it, or open Query Properties and set an automatic refresh interval, say every 10 minutes for data that changes during the day. And if point-and-click doesn't capture what you need, the `Html.Table` M function lets you extract elements by CSS selector, for example `Html.Table(html, {{"Price", ".product-price"}})`. That covers most of what Google Sheets' IMPORTXML can do, inside Excel.

The failure mode is simple: Power Query fetches raw HTML. Any table assembled by JavaScript — most modern product pages, dashboards, and search results — arrives empty. No amount of retrying fixes that; the content simply isn't in the HTML that Excel receives. Same story for sites that show a CAPTCHA or a bot-detection page instead of the table you saw in your browser.

When you hit that wall, you have two options: write code yourself, or route the request through a service that handles the hard parts.

## Method 2: VBA — more control, same JavaScript ceiling

VBA can download a page with `MSXML2.XMLHTTP.6.0`, load the response into an HTML document object, walk the DOM with `getElementsByTagName("table")`, and write each cell's `innerText` into a sheet array. Modern approaches don't need Internet Explorer at all.

It's free and it stays entirely inside Excel, which matters if you're building something for colleagues who just open a workbook. But you're writing and maintaining the parser, and the ceiling is the same as Power Query's: `XMLHTTP` doesn't execute JavaScript, doesn't rotate IPs, and won't solve a CAPTCHA. For a handful of stable internal pages it's fine. For anything guarded, it's a dead end.

## Method 3: A scraping API — and how the data lands in Excel

A scraping API sits between Excel and the target site. You send the URL, the service handles proxies, headers, CAPTCHAs, and JavaScript rendering, then returns structured data you can import. This is the route that keeps working when the first two methods return blank cells.

ScraperAPI is one of the longer-running services here, and its workflow maps cleanly onto spreadsheet work:

**Direct API calls.** A GET request to the API endpoint with your URL and API key returns the page. You can set `outputFormat` to `csv` or `json`, and with `autoparse` enabled on supported sites, you get ready-parsed structured data — no parsing code on your side. From there it's a straight import into Excel. The service also works as a proxy port, so existing scripts can point at it without restructuring. One detail from the docs worth following: set a 70-second timeout in your application, especially for hard-to-scrape domains.

**DataPipeline — the no-code route.** If you'd rather not touch code at all, ScraperAPI's dashboard includes a batch tool:

1. Open **DataPipeline** in the sidebar and click **New project**
2. Choose **Extract data from arbitrary websites**
3. Under input settings, enter URLs one per line, or upload a `.txt`/`.csv` file with up to 100,000 URLs
4. Choose output as a dashboard download (or a Webhook URL for larger setups)
5. Pick a frequency — a single run, or a Daily/Weekly schedule
6. Review the job summary, which shows the estimated credit cost before anything runs

A few practical details live in the fine print: recurring schedules require a paid plan, results are retained for 30 days so you need to download them before then, and options like premium residential/mobile IPs, JavaScript rendering, country targeting, and 404 retries are toggles in the project settings rather than code changes. For feeding an Excel file on a schedule, "upload a CSV of 10,000 product URLs, scrape weekly, download the results" is the entire workflow.

## ScraperAPI plans and current pricing

Here's the full plan lineup from the official pricing page. Every paid plan includes JavaScript rendering, premium residential and mobile IPs, advanced anti-bot bypassing, structured data APIs, DataPipeline access, and a 99.9% uptime guarantee. Annual billing takes 10% off every tier.

| Plan | Monthly price | Annual (per month) | Credits | Threads | Geo targeting |
| --- | --- | --- | --- | --- | --- |
| Free | $0 | — | 1,000/mo | 5 | US & EU |
| Hobby | $49 | $44.10 | 100,000 | 20 | US & EU only |
| Startup | $149 | $134.10 | 1,000,000 | 50 | US & EU only |
| Business | $299 | $269.10 | 3,000,000 | 100 | Country-level |
| Scaling | $475 | $427.50 | 5,000,000 | 200 | Country-level |
| Professional | $975 | $877.50 | 10,500,000 | 300 | Country-level |
| Advanced | $1,975 | $1,777.50 | 21,500,000 | 500 | Country-level |
| Enterprise | Custom | Custom | 22,000,000+ | 500+ | Country-level |
| Plan | Get started |  |  |  |  |
| --- | --- |  |  |  |  |
| Free | [Create a free account](https://www.scraperapi.com/?fp_ref=coupons) |  |  |  |  |
| Hobby | [Check the Hobby plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Startup | [Check the Startup plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Business | [Check the Business plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Scaling | [Check the Scaling plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Professional | [Check the Professional plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Advanced | [Check the Advanced plan](https://www.scraperapi.com/pricing/?fp_ref=coupons) |  |  |  |  |
| Enterprise | [Talk to sales about Enterprise](https://www.scraperapi.com/contact-sales/?fp_ref=coupons) |  |  |  |  |

New accounts also get a 7-day trial with 5,000 free API credits on top of the permanent free tier, so you can run a real test against your target site before the first invoice.

## The credit math nobody explains: why 100,000 credits ≠ 100,000 pages

This is the part most comparisons gloss over, and it decides whether a plan is enough for your project. One credit equals a standard page request — but some sites cost far more:

| Target | Credits per request |
| --- | --- |
| Standard page | 1 |
| Amazon | 5 |
| Google / Bing (incl. subdomains) | 25 |
| LinkedIn | 30 |
| Sites behind Cloudflare, Datadome, PerimeterX (when bypassed) | +10 |

So a Hobby plan's 100,000 credits is 100,000 ordinary pages, but only about 4,000 Google searches or roughly 3,300 LinkedIn requests. If your "web scraping to Excel" project is actually "pull Google results," do this math before choosing a tier.

Two safeguards make this manageable. The dashboard's Domain Cost Estimator shows the exact cost for any URL before you burn credits, and you can set a `max_cost` per scrape so a single request can never exceed your limit. If a request goes over your concurrent thread limit, the API returns a 429 and asks you to retry — there's no extra charge for it.

## Billing quirks worth knowing before you subscribe

Several details only appear in the billing FAQ, and each one changes how you'd use a plan:

- **Credits running out pauses scraping by default.** If you want uninterrupted service, auto-upgrade (opt-in, available on Hobby through Professional) bumps you to the next tier automatically. Pay-As-You-Go — continuing at a fixed per-credit rate — is only available on Scaling, Professional, Advanced, and Enterprise.
- **Pay-As-You-Go has a spending cap.** You set a monthly limit and service stops at it. It's a cap, not a pre-charge, so there's no surprise-bill mechanics to worry about.
- **Annual plans can renew early** for a fresh batch of credits if you burn through the year's allocation.
- **Cancellation is free anytime**, and there's a formal policy: "We offer a 7-day no questions asked refund policy." Combined with the free trial, you have two separate ways to walk away without cost.
- Billing is by request count, not bandwidth, which is why a 50MB PDF download and a 40KB product page both count as one request (assuming standard credit cost).

For a typical Excel workflow — a few thousand page pulls per month — Hobby is usually the right entry point. Startup's 10x credit jump at 3x the price only makes sense if your list is genuinely that long or you're scraping expensive domains. Given the 7-day refund window, buying a month of Hobby to validate your target site is a low-risk way to find out.

👉 [See current Hobby and Startup pricing](https://www.scraperapi.com/pricing/?fp_ref=coupons)

## How ScraperAPI compares on price, and when something else fits

For context on the paid landscape: Apify starts around $29/month, Oxylabs around $30/month, Zyte around $100/month, and Bright Data around $499/month. Only ScraperAPI and Apify offer a free plan in that set. Exact configurations vary enough between providers that these starting prices say more about entry points than value — Bright Data's floor is aimed at a different scale of buyer entirely.

The honest fit assessment: if your targets are static HTML tables, none of these are needed — Power Query does it for free. If you need scheduled extraction of JavaScript-heavy or protected sites into spreadsheets and don't want to maintain infrastructure, ScraperAPI's middle tiers are priced for exactly that, and the free trial lets you verify success rates on your specific targets first. If your volumes are enormous or highly specialized, any of the enterprise-focused providers deserve a look on their own merits.

## FAQ

**Do I need a credit card for the ScraperAPI trial?**
The trial gives you 5,000 API credits over your first 7 days; third-party reviews report that no credit card is typically required to start it. The free tier's 1,000 monthly credits definitely require nothing but an account.

**Can Excel refresh scraped data automatically?**
With Power Query, yes — set a refresh interval in Query Properties for static-HTML sources. For protected or JS-rendered sites, schedule DataPipeline jobs (Daily/Weekly on a paid plan) and download the output on each run.

**What happens if I exceed my plan's limits?**
Over the thread limit, requests return a 429 with no charge — just retry. Over the credit limit, scraping pauses until renewal unless you've enabled auto-upgrade or (on Scaling and above) Pay-As-You-Go.

**Is there a refund if it doesn't work for my target site?**
Yes. ScraperAPI offers a 7-day no-questions-asked refund on top of the free trial, so a mismatch between marketing and reality on a difficult domain costs you nothing.

## Which route should you actually take

Match the method to the target site, and the decision mostly makes itself:

- **Static HTML table, occasional pulls** → Power Query. Free, refreshable, done in minutes.
- **Static pages, custom logic, everything must live in one workbook** → VBA.
- **JavaScript-rendered or protected sites, or scheduled bulk lists** → a scraping API. Test on the free tier first, then Hobby at $49/month covers 100,000 standard pages — fewer if Google or LinkedIn are in the mix, so check the Domain Cost Estimator before picking a tier.
- **Millions of requests per month** → Enterprise is custom-priced; contact sales rather than guessing.

One last check before subscribing: run two or three of your real target URLs through the trial and confirm the output format works for your spreadsheet import. Success rates on a difficult domain are the single biggest variable in whether this route saves you time or adds a maintenance burden — and it's the one thing you can verify before money changes hands.

👉 [Test your target sites with the free trial](https://www.scraperapi.com/?fp_ref=coupons)
