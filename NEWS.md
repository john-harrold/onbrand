# onbrand 1.0.3

## Table formatting

* Added functions to take a large table and create smaller tables to span across multiple pages in reports (`span_table()` and `build_span()`)

* Render markdown in specified part of a flext table (`ft_apply_md()`)

# onbrand 1.0.2

## Changes to yaml file format

* The addition of notes option for tables and figures in Word reporting requires the `Notes` document default be added to any `report.yaml` files:

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

* Adding `post_processing` option for `rdocx` and `rpptx` sections. These can
  be omitted or set to NULL and they will have no effect. Any R code here is
  evaluated just before saving (when running `save_report()`) and you can
  modify the object `rpt` (the officer report object from the `obnd` object).

```
rdocx:
  post_processing: NULL
rpptx:
  post_processing: NULL
```

## Minor Improvements

* Added `fig_start_at` and `tab_start_at` arguments to `report_add_doc_content()` to support chapter specific numbering (e.g. Figure 3-1, Figure 3-2, Figure 4-1, etc).

* Added notes option for tables and figures in Word reporting.

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
