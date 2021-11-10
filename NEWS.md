# onbrand 1.0.2

## Changes to yaml file format

* Added notes option for tables and figures in Word reporting. This will requires the `Notes` document default be added to any `report.yaml` files:

```
doc_def:                     
  Notes: Notes
```

* Added the formatting options for separator (regional configuration variable that seems to be different between Europe and the US), and several table and figure formatting options: order, sequence IDs, and how numbering will be done (allows Figure 1 or Figure 3-1 (for section 3), etc.)

```
formatting:
  separator:                 ","
  Table_Order:
    - table
    - notes
    - caption
  Figure_Order:
    - figure
    - notes
    - caption
  Figure_Seq_Id:             "Figure"
  Figure_Number: |-
    list(officer::run_autonum(pre_label  = "", 
                              seq_id     = Caption_Seq_Id, 
                              post_label = "", 
                              start_at   = Caption_Start_At))
  Table_Seq_Id:              "Table"
  Table_Number: |-
    list(officer::run_autonum(pre_label  = "", 
                              seq_id     = Caption_Seq_Id, 
                              post_label = "", 
                              start_at   = Caption_Start_At))
```

* Added `fig_start_at` and `tab_start_at` arguments to `report_add_doc_content()` to support this.

## Minor Improvements

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
