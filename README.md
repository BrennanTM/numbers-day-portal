# VETS Numbers Day portal

A single-page tool that builds the monthly Numbers Day email (veteran unemployment rates after each
BLS Employment Situation release) from the BLS Public Data API, entirely in the browser, and copies
it formatted for Outlook. Open `index.html` or use the GitHub Pages site for this repository.

Setup: paste a BLS API registration key once (saved in the browser only). Drop the monthly BLS
disability package PDF on the page and it reads Table 16 itself, or type the four rates, or
import a CSV. Press Build, then "Copy email for Outlook".

Releases carry `numbersday.pyz`, the command-line companion for Windows machines with Python:
it writes the Word report, reads the same PDF, and can open an Outlook draft.

No key and no data live in this repository. Data: U.S. Bureau of Labor Statistics, Current
Population Survey, retrieved live from api.bls.gov.
