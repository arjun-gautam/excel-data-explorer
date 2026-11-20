# Data Explorer

**Live site: https://excel-data-explorer.web.app**

Upload an .xlsx file. The app reads the column names, works out each column's type (number, date, yes/no, text), and builds filters and sorting that fit those types. No formulas needed. Download the result as CSV.

    npm install
    npm run dev      # development
    npm run build    # makes ONE file: dist/index.html (double-click to open, no server needed)

The file is read in the browser only. Nothing is uploaded or stored.
