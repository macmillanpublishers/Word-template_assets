# Word-template_assets
Macmillan-specific assets for our Word tools macros

## Styles_Bookmaker.csv
Single-column list of Macmillan styles that are available in Bookmaker for PDFs. Used by the Bookmaker Check macro. The macro downloads this file each time it is run, so users do not need to get a formal update to use any additions to this list. Any changes committed to `master` are available to the user within 5 minutes (via Confluence-Git connector).

## Styles_Mapping.json
Maps old Macmillan template style names to new versions. Not currently live but will be added to Cleanup and Character Styles macros. Consists of two objects: `"renamed"` and `"removed"`; both contain `"old style name":"new style name"` pairs.

## headings.json
Keys are Macmillan styles that are acceptable section headings. Value is always "False". Used in vba_utilities/genUtils/Reports.bas

## sections.json
object keys are the FIRST word of Macmillan style-name sections. Each contains an aobject with the heading style that should precede a style in that section, as well as the text of that heading.
