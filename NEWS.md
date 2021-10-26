# onbrand 1.0.2

## Minor Improvements

* Added notes option for tables and figures in Word reporting. This will requires the `Notes` document default be added to any `report.yaml` files:

```
doc_def:                     
  Notes: Notes
```

* Created `ftext` formatting for captions and notes added to figures and tables.

* Support for multipage figures and tables (e.g. pagenated figures). By specifying the same key for a figure (or table) the first instance will be the figure and subsequent instances will be references to the first. 

* Support for crossreferencing figures and tables with markdown.


# onbrand 1.0.1       

## Minor Improvements

* updated diagnostic messages to include package name

* added NEWS.md

* Added the function `template_details` 
  * Provide onbrand names for template elements
  * Created tests
  * Updated vignettes 

# onbrand 1.0.0 

* Submitted to CRAN
