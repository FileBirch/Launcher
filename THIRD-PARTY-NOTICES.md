# Distribution audit

The shipped runtime uses first-party ES modules, browser APIs (including Compression/DecompressionStream, DOMParser, Web Crypto and Canvas), and system fonts. The bounded ZIP reader/writer and spreadsheet parsers are local source; no JSZip, SheetJS, pako, external font or third-party runtime library is bundled. The XLSX fixture is test data. Playwright and openpyxl, when installed for tests, are test-only and are not redistributed. Microsoft/Excel/Google references identify file-format compatibility, not bundled software. No separately attributed copied snippet was identified. Branding/artwork are excluded from MIT.

This is an inventory of identified shipping components, not a license grant for excluded material. Preserve each applicable notice when repackaging.
