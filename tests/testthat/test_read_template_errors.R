# Helper to write a temporary yaml file
write_tmp_yaml = function(yaml_list) {
  f = tempfile(fileext = ".yaml")
  yaml::write_yaml(yaml_list, f)
  f
}

pptx_template = file.path(system.file(package="onbrand"), "templates", "report.pptx")
docx_template = file.path(system.file(package="onbrand"), "templates", "report.docx")

# --- Mapping file errors ---

test_that("read_template: mapping file not found", {
  expect_error(
    read_template(template = pptx_template,
                  mapping  = "/nonexistent/path.yaml",
                  verbose  = FALSE))
})

test_that("read_template: PowerPoint template with missing rpptx section", {
  yaml_file = write_tmp_yaml(list(rdocx = list()))
  expect_error(
    read_template(template = pptx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word template with missing rdocx section", {
  yaml_file = write_tmp_yaml(list(rpptx = list()))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

# --- PowerPoint validation errors ---

test_that("read_template: PowerPoint layout not found in template", {
  yaml_file = write_tmp_yaml(list(
    rpptx = list(
      master = "Office Theme",
      templates = list(
        fake_layout = list(
          title = list(ph_label = "Title 1", content_type = "text")
        )
      ),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = TRUE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Helvetica", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent")
      )
    )
  ))
  expect_error(
    read_template(template = pptx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: PowerPoint placeholder not found", {
  yaml_file = write_tmp_yaml(list(
    rpptx = list(
      master = "Office Theme",
      templates = list(
        title_slide = list(
          title = list(ph_label = "NONEXISTENT PLACEHOLDER", content_type = "text")
        )
      ),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = TRUE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Helvetica", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent")
      )
    )
  ))
  expect_error(
    read_template(template = pptx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: PowerPoint missing md_def section", {
  yaml_file = write_tmp_yaml(list(
    rpptx = list(
      master = "Office Theme",
      templates = list(
        title_slide = list(
          title = list(ph_label = "Title 1", content_type = "text"),
          sub_title = list(ph_label = "Subtitle 2", content_type = "text")
        )
      )
    )
  ))
  expect_error(
    read_template(template = pptx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: PowerPoint missing required style in md_def", {
  yaml_file = write_tmp_yaml(list(
    rpptx = list(
      master = "Office Theme",
      templates = list(
        title_slide = list(
          title = list(ph_label = "Title 1", content_type = "text"),
          sub_title = list(ph_label = "Subtitle 2", content_type = "text")
        )
      ),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = TRUE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Helvetica", vertical.align = "baseline",
                       shading.color = "transparent")
        # Missing Table_Labels
      )
    )
  ))
  expect_error(
    read_template(template = pptx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

# --- Word validation errors ---

test_that("read_template: Word style not found in template", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Fake_Style = "Nonexistent Style Name"),
      doc_def = list(Text = "Normal", Table = "Normal",
                     Table_Caption = "Normal", Figure_Caption = "Normal",
                     Notes = "Normal"),
      formatting = list(separator = ","),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12),
        Fake_Style = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word missing doc_def entries", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal"),
      doc_def = list(Text = "Normal"),
      # Missing Table, Table_Caption, Figure_Caption, Notes
      formatting = list(separator = ","),
      md_def = list(
        default = list(color = "black", font.size = 12),
        Table_Labels = list(color = "black", font.size = 12),
        Normal = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word doc_def references undefined style", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal"),
      doc_def = list(Text = "Normal", Table = "Undefined_Style",
                     Table_Caption = "Normal", Figure_Caption = "Normal",
                     Notes = "Normal"),
      formatting = list(separator = ","),
      md_def = list(
        default = list(color = "black", font.size = 12),
        Table_Labels = list(color = "black", font.size = 12),
        Normal = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word missing separator in formatting", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Notes = "Notes",
                    Table = "Table Grid",
                    Table_Caption = "table title",
                    Figure_Caption = "graphic title"),
      doc_def = list(Text = "Normal", Table = "Table",
                     Table_Caption = "Table_Caption",
                     Figure_Caption = "Figure_Caption",
                     Notes = "Notes"),
      formatting = list(),
      # Missing separator
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12),
        Notes = list(color = "black", font.size = 10),
        Table = list(color = "black", font.size = 12),
        Table_Caption = list(color = "black", font.size = 12),
        Figure_Caption = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word invalid separator value", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Notes = "Notes",
                    Table = "Table Grid",
                    Table_Caption = "table title",
                    Figure_Caption = "graphic title"),
      doc_def = list(Text = "Normal", Table = "Table",
                     Table_Caption = "Table_Caption",
                     Figure_Caption = "Figure_Caption",
                     Notes = "Notes"),
      formatting = list(separator = "|"),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12),
        Notes = list(color = "black", font.size = 10),
        Table = list(color = "black", font.size = 12),
        Table_Caption = list(color = "black", font.size = 12),
        Figure_Caption = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word style missing from md_def", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Notes = "Notes",
                    Table = "Table Grid",
                    Table_Caption = "table title",
                    Figure_Caption = "graphic title"),
      doc_def = list(Text = "Normal", Table = "Table",
                     Table_Caption = "Table_Caption",
                     Figure_Caption = "Figure_Caption",
                     Notes = "Notes"),
      formatting = list(separator = ","),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12)
        # Missing Notes, Table, Table_Caption, Figure_Caption
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word invalid Table_Order", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Notes = "Notes",
                    Table = "Table Grid",
                    Table_Caption = "table title",
                    Figure_Caption = "graphic title"),
      doc_def = list(Text = "Normal", Table = "Table",
                     Table_Caption = "Table_Caption",
                     Figure_Caption = "Figure_Caption",
                     Notes = "Notes"),
      formatting = list(separator = ",",
                        Table_Order = list("table", "invalid_element"),
                        Figure_Order = list("figure", "notes", "caption")),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12),
        Notes = list(color = "black", font.size = 10),
        Table = list(color = "black", font.size = 12),
        Table_Caption = list(color = "black", font.size = 12),
        Figure_Caption = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})

test_that("read_template: Word invalid Figure_Order", {
  yaml_file = write_tmp_yaml(list(
    rdocx = list(
      styles = list(Normal = "Normal", Notes = "Notes",
                    Table = "Table Grid",
                    Table_Caption = "table title",
                    Figure_Caption = "graphic title"),
      doc_def = list(Text = "Normal", Table = "Table",
                     Table_Caption = "Table_Caption",
                     Figure_Caption = "Figure_Caption",
                     Notes = "Notes"),
      formatting = list(separator = ",",
                        Table_Order = list("table", "notes", "caption"),
                        Figure_Order = list("figure", "bad_element")),
      md_def = list(
        default = list(color = "black", font.size = 12, bold = FALSE,
                       italic = FALSE, underlined = FALSE,
                       font.family = "Cambria (Body)", vertical.align = "baseline",
                       shading.color = "transparent"),
        Table_Labels = list(color = "black", font.size = 12, bold = TRUE,
                            italic = FALSE, underlined = FALSE,
                            font.family = "Helvetica", vertical.align = "baseline",
                            shading.color = "transparent"),
        Normal = list(color = "black", font.size = 12),
        Notes = list(color = "black", font.size = 10),
        Table = list(color = "black", font.size = 12),
        Table_Caption = list(color = "black", font.size = 12),
        Figure_Caption = list(color = "black", font.size = 12)
      )
    )
  ))
  expect_error(
    read_template(template = docx_template,
                  mapping  = yaml_file,
                  verbose  = FALSE))
})
