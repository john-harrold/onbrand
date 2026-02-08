library(onbrand)
library(officer)

# Default format for comparisons
default_props_str <- 'officer::fp_text(bold = FALSE,font.size = 12,italic = FALSE,underlined = FALSE,color = "black",shading.color = "transparent",vertical.align = "baseline",font.family = "Cambria (Body)")'

test_that("Plain text without markdown", {
  result <- md_to_officer("Plain text")

  expect_equal(length(result), 1)
  expect_true("pgraph_1" %in% names(result))
  expect_equal(result$pgraph_1$pele$p_1$text, "Plain text")
  expect_equal(result$pgraph_1$pele$p_1$md_name, "none")

  # Check ftext_cmd
  expected_ftext <- paste0('officer::ftext("Plain text", prop=', default_props_str, ')')
  expect_equal(result$pgraph_1$ftext_cmd, expected_ftext)

  # Check fpar_cmd
  expected_fpar <- paste0("officer::fpar(", expected_ftext, ")")
  expect_equal(result$pgraph_1$fpar_cmd, expected_fpar)
})

test_that("Bold text with double asterisks", {
  result <- md_to_officer("Be **bold**!")

  expect_equal(length(result$pgraph_1$pele), 3)
  expect_equal(result$pgraph_1$pele$p_1$text, "Be ")
  expect_equal(result$pgraph_1$pele$p_2$text, "bold")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "bold")
  expect_equal(result$pgraph_1$pele$p_3$text, "!")

  # Check that bold property is present in ftext_cmd
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("Be ", result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold", result$pgraph_1$ftext_cmd))
})

test_that("Bold text with double underscores", {
  result <- md_to_officer("Be __bold__!")

  expect_equal(result$pgraph_1$pele$p_2$text, "bold")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "bold")

  # Check that bold property is present
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Italic text with single asterisk", {
  result <- md_to_officer("This is *italic* text")

  expect_equal(result$pgraph_1$pele$p_2$text, "italic")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "italic_st")

  # Check that italic property is present
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Italic text with single underscore", {
  result <- md_to_officer("This is _italic_ text")

  expect_equal(result$pgraph_1$pele$p_2$text, "italic")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "italic_us")

  # Check that italic property is present
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Subscript text", {
  result <- md_to_officer("H~2~O")

  expect_equal(result$pgraph_1$pele$p_2$text, "2")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "subscript")

  # Check that subscript property is present
  expect_true(grepl('vertical.align = "subscript"', result$pgraph_1$ftext_cmd))
})

test_that("Superscript text", {
  result <- md_to_officer("E=mc^2^")

  expect_equal(result$pgraph_1$pele$p_2$text, "2")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "superscript")

  # Check that superscript property is present
  expect_true(grepl('vertical.align = "superscript"', result$pgraph_1$ftext_cmd))
})

test_that("Colored text", {
  result <- md_to_officer("This is <color:red>red text</color>")

  expect_equal(result$pgraph_1$pele$p_2$text, "red text")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "color")

  # Check that color property is present
  expect_true(grepl('color = "red"', result$pgraph_1$ftext_cmd))
})

test_that("Shaded text", {
  result <- md_to_officer("This is <shade:#33ff33>shaded text</shade>")

  expect_equal(result$pgraph_1$pele$p_2$text, "shaded text")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "shading_color")

  # Check that shading color property is present
  expect_true(grepl('shading.color = "#33ff33"', result$pgraph_1$ftext_cmd))
})

test_that("Font family change", {
  result <- md_to_officer("This uses <ff:symbol>symbol font</ff>")

  expect_equal(result$pgraph_1$pele$p_2$text, "symbol font")
  expect_equal(result$pgraph_1$pele$p_2$md_name, "font_family")

  # Check that font family property is present
  expect_true(grepl('font.family = "symbol"', result$pgraph_1$ftext_cmd))
})

test_that("Reference markdown", {
  result <- md_to_officer("See <ref:table1> for details")

  expect_equal(result$pgraph_1$pele$p_2$text, 'officer::run_reference("table1")')
  expect_equal(result$pgraph_1$pele$p_2$md_name, "reference")

  # Check that run_reference is in the command
  expect_true(grepl('officer::run_reference\\("table1"\\)', result$pgraph_1$ftext_cmd))
})

test_that("Multiple paragraphs separated by double newlines", {
  result <- md_to_officer("First paragraph\n\nSecond paragraph")

  expect_equal(length(result), 2)
  expect_true("pgraph_1" %in% names(result))
  expect_true("pgraph_2" %in% names(result))
  expect_equal(result$pgraph_1$pele$p_1$text, "First paragraph")
  expect_equal(result$pgraph_2$pele$p_1$text, "Second paragraph")
})

test_that("Multiple paragraphs with different line endings", {
  # Test with \r\n (Windows)
  result <- md_to_officer("First\r\n\r\nSecond")
  expect_equal(length(result), 2)

  # Test with \r (old Mac)
  result <- md_to_officer("First\r\rSecond")
  expect_equal(length(result), 2)
})

test_that("Nested formatting: bold and italic (asterisk)", {
  result <- md_to_officer("**_bold and italic_**")

  # Should have bold as outer and italic as inner
  expect_equal(result$pgraph_1$pele$p_1$text, "bold and italic")

  # Check that both properties are present
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: bold and italic (alternative syntax)", {
  result <- md_to_officer("***bold and italic***")

  # Check that both properties are present
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: colored bold text", {
  result <- md_to_officer("<color:blue>**bold blue**</color>")

  expect_equal(result$pgraph_1$pele$p_1$text, "bold blue")

  # Check that both properties are present
  expect_true(grepl('color = "blue"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: colored italic text", {
  result <- md_to_officer("<color:red>*italic red*</color>")

  expect_equal(result$pgraph_1$pele$p_1$text, "italic red")

  # Check that both properties are present
  expect_true(grepl('color = "red"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: bold italic and colored", {
  result <- md_to_officer("<color:green>**_bold italic green_**</color>")

  expect_equal(result$pgraph_1$pele$p_1$text, "bold italic green")

  # Check that all three properties are present
  expect_true(grepl('color = "green"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: shaded bold text", {
  result <- md_to_officer("<shade:#ffff00>**bold yellow background**</shade>")

  expect_equal(result$pgraph_1$pele$p_1$text, "bold yellow background")

  # Check that both properties are present
  expect_true(grepl('shading.color = "#ffff00"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: colored subscript", {
  result <- md_to_officer("<color:blue>H~2~O</color>")

  # Check that color and subscript are present
  expect_true(grepl('color = "blue"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('vertical.align = "subscript"', result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: colored superscript", {
  result <- md_to_officer("<color:red>E=mc^2^</color>")

  # Check that color and superscript are present
  expect_true(grepl('color = "red"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('vertical.align = "superscript"', result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: font family with bold", {
  result <- md_to_officer("<ff:Arial>**bold Arial**</ff>")

  expect_equal(result$pgraph_1$pele$p_1$text, "bold Arial")

  # Check that both properties are present
  expect_true(grepl('font.family = "Arial"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: font family with italic", {
  result <- md_to_officer("<ff:Courier>*italic Courier*</ff>")

  expect_equal(result$pgraph_1$pele$p_1$text, "italic Courier")

  # Check that both properties are present
  expect_true(grepl('font.family = "Courier"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: color and shade combined", {
  result <- md_to_officer("<shade:#cccccc><color:red>red on gray</color></shade>")

  expect_equal(result$pgraph_1$pele$p_1$text, "red on gray")

  # Check that both properties are present
  expect_true(grepl('color = "red"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('shading.color = "#cccccc"', result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: bold within colored within shaded", {
  result <- md_to_officer("<shade:#ffffcc><color:blue>**triple nested**</color></shade>")

  expect_equal(result$pgraph_1$pele$p_1$text, "triple nested")

  # Check that all three properties are present
  expect_true(grepl('shading.color = "#ffffcc"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('color = "blue"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: multiple nested elements in sequence", {
  result <- md_to_officer("Start <color:red>**bold red**</color> then <color:blue>*italic blue*</color> end")

  # Should have multiple elements
  expect_true(length(result$pgraph_1$pele) >= 4)

  # Check for red color and bold
  expect_true(grepl('color = "red"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))

  # Check for blue color and italic
  expect_true(grepl('color = "blue"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Nested formatting: bold italic with font family and color", {
  result <- md_to_officer("<ff:Times><color:purple>**_everything_**</color></ff>")

  expect_equal(result$pgraph_1$pele$p_1$text, "everything")

  # Check that all four properties are present
  expect_true(grepl('font.family = "Times"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('color = "purple"', result$pgraph_1$ftext_cmd))
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

test_that("Mixed formatted and plain text", {
  result <- md_to_officer("Start **bold** middle *italic* end")

  expect_equal(length(result$pgraph_1$pele), 5)
  expect_equal(result$pgraph_1$pele$p_1$text, "Start ")
  expect_equal(result$pgraph_1$pele$p_2$text, "bold")
  expect_equal(result$pgraph_1$pele$p_3$text, " middle ")
  expect_equal(result$pgraph_1$pele$p_4$text, "italic")
  expect_equal(result$pgraph_1$pele$p_5$text, " end")
})

test_that("Complex example with multiple formatting", {
  result <- md_to_officer("The formula H~2~O equals water, and E=mc^2^ is Einstein's equation.")

  # Check that both subscript and superscript are present
  expect_true(grepl('vertical.align = "subscript"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('vertical.align = "superscript"', result$pgraph_1$ftext_cmd))
})

test_that("Custom default format", {
  custom_format <- list(
    color          = "blue",
    font.size      = 14,
    bold           = TRUE,
    italic         = FALSE,
    underlined     = FALSE,
    font.family    = "Arial",
    vertical.align = "baseline",
    shading.color  = "transparent"
  )

  result <- md_to_officer("Plain text", default_format = custom_format)

  # Check that custom defaults are in the command
  expect_true(grepl('color = "blue"', result$pgraph_1$ftext_cmd))
  expect_true(grepl('font.size = 14', result$pgraph_1$ftext_cmd))
  expect_true(grepl('bold = TRUE', result$pgraph_1$ftext_cmd))
  expect_true(grepl('font.family = "Arial"', result$pgraph_1$ftext_cmd))
})

test_that("fpar_cmd can be evaluated", {
  result <- md_to_officer("Be **bold**!")

  # Test that the fpar_cmd can be evaluated without error
  expect_no_error({
    fpar_obj <- eval(parse(text = result$pgraph_1$fpar_cmd))
  })

  # Check that the result is an fpar object
  fpar_obj <- eval(parse(text = result$pgraph_1$fpar_cmd))
  expect_true(inherits(fpar_obj, "fpar"))
})

test_that("as_paragraph_cmd can be evaluated", {
  result <- md_to_officer("Be **bold**!")

  # Test that the as_paragraph_cmd can be evaluated without error
  expect_no_error({
    as_para_obj <- eval(parse(text = result$pgraph_1$as_paragraph_cmd))
  })
})

test_that("Special characters in text", {
  result <- md_to_officer('Text with "quotes" and \'apostrophes\'')

  expect_equal(result$pgraph_1$pele$p_1$text, 'Text with "quotes" and \'apostrophes\'')
  expect_true(grepl("quotes", result$pgraph_1$ftext_cmd))
})

test_that("Markdown at start and end of string", {
  result <- md_to_officer("**bold at start** and **bold at end**")

  expect_equal(result$pgraph_1$pele$p_1$text, "bold at start")
  expect_equal(result$pgraph_1$pele$p_1$md_name, "bold")
  expect_equal(result$pgraph_1$pele$p_3$text, "bold at end")
  expect_equal(result$pgraph_1$pele$p_3$md_name, "bold")
})

test_that("Consecutive markdown elements", {
  result <- md_to_officer("**bold***italic*")

  # Should parse as two separate formatted elements
  expect_true(grepl("bold = TRUE", result$pgraph_1$ftext_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$ftext_cmd))
})

# Bug report: https://github.com/john-harrold/onbrand/issues/13
test_that("Bug #13: Italic text nested inside bold text - FIXED", {
  # This bug has been fixed! Bold text containing italic text now properly
  # preserves the bold-only portions before and after the italic section
  result <- md_to_officer("**Be bold and *italic* **: ")

  # Fixed behavior: correctly creates 4 elements
  expect_equal(length(result$pgraph_1$pele), 4)
  expect_equal(result$pgraph_1$pele$p_1$text, "Be bold and ")
  expect_true(grepl("bold = TRUE", result$pgraph_1$pele$p_1$props_cmd))
  expect_false(grepl("italic = TRUE", result$pgraph_1$pele$p_1$props_cmd))

  expect_equal(result$pgraph_1$pele$p_2$text, "italic")
  expect_true(grepl("bold = TRUE", result$pgraph_1$pele$p_2$props_cmd))
  expect_true(grepl("italic = TRUE", result$pgraph_1$pele$p_2$props_cmd))

  expect_equal(result$pgraph_1$pele$p_3$text, " ")
  expect_true(grepl("bold = TRUE", result$pgraph_1$pele$p_3$props_cmd))
  expect_false(grepl("italic = TRUE", result$pgraph_1$pele$p_3$props_cmd))

  expect_equal(result$pgraph_1$pele$p_4$text, ": ")
  expect_false(grepl("bold = TRUE", result$pgraph_1$pele$p_4$props_cmd))
  expect_false(grepl("italic = TRUE", result$pgraph_1$pele$p_4$props_cmd))
})
