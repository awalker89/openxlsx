


context("Style Parsing")



test_that("parsing border xml", {
  wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
  styles <- getStyles(wb = wb)


  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, "medium",
    "medium", "medium", "medium", NULL, NULL, NULL, NULL, NULL,
    NULL, NULL, NULL, NULL, "thin", NULL, "thin", "thin", NULL,
    "thin", "thin", "thin", "thin", "thin", "thin", "thin", NULL,
    "thin", "thin", "medium", "medium", "medium", "medium", "thin",
    "medium", "medium", "thin", NULL, "medium", "medium", "medium",
    "thin", "thin", "medium", "medium", "thin", "thin", "thick",
    NULL, "thick", "thick", "thick", NULL, NULL, NULL, NULL,
    NULL, "medium", "medium", NULL, "medium", "mediumDashed",
    "mediumDashed", "mediumDashed", NULL, NULL, NULL, NULL, NULL,
    NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderBottom"))


  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,
    NULL, NULL, NULL, NULL, NULL, "thin", "thin", "thin", NULL,
    NULL, NULL, NULL, "medium", NULL, NULL, NULL, NULL, NULL,
    "thin", NULL, "thin", "thick", NULL, "medium", "thin", "thin",
    "thin", "thin", "thick", "thick", "thin", "thin", "thin",
    "medium", "medium", "thin", "thick", "thick", "medium", "thin",
    "thick", "thick", "medium", "thin", "thin", "medium", "thin",
    "thin", "thin", "medium", "medium", "medium", NULL, NULL,
    NULL, NULL, NULL, NULL, "mediumDashed", "mediumDashed", "mediumDashed",
    NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderTop"))



  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, "medium", NULL, "medium",
    NULL, NULL, NULL, NULL, NULL, NULL, "thin", NULL, NULL, "thin",
    NULL, NULL, "thin", "medium", NULL, NULL, NULL, NULL, "thin",
    "thin", "thin", NULL, "thin", NULL, NULL, NULL, "thin", "medium",
    "thin", "thin", "thin", "thin", "medium", "thin", "thin",
    NULL, "thin", "thick", "thin", "thick", "thick", "thin",
    "thin", "thin", "thin", "thick", NULL, "thin", "thin", "thin",
    "medium", NULL, NULL, "medium", NULL, "medium", NULL, "medium",
    NULL, "mediumDashed", NULL, NULL, NULL, NULL, NULL, NULL,
    NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderLeft"))


  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, "medium", NULL, "medium",
    NULL, NULL, NULL, "medium", NULL, NULL, NULL, NULL, NULL,
    "thin", NULL, NULL, "thin", NULL, NULL, NULL, "thin", NULL,
    "thin", NULL, "thin", "thin", "thin", "thin", "thick", NULL,
    "thick", "medium", "thin", "thin", "thin", "thin", "medium",
    "thin", "thin", "medium", NULL, "medium", "thin", "thin",
    "medium", "medium", "thin", "thick", "medium", "medium",
    "thin", "medium", "thin", NULL, "thick", NULL, NULL, "medium",
    NULL, "medium", NULL, NULL, NULL, "medium", NULL, NULL, "mediumDashed",
    NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderRight"))



  ## COLOURS
  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), NULL, NULL, NULL,
    NULL, NULL, NULL, NULL, NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, structure(list(theme = "6"), .Names = "theme"), structure(list(
      theme = "6"
    ), .Names = "theme"), NULL, structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), NULL, structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "3"
    ), .Names = "theme"), structure(list(theme = "3"), .Names = "theme"),
    structure(list(theme = "3"), .Names = "theme"), structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(theme = "6"), .Names = "theme"), structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure(list(theme = "6"), .Names = "theme"), structure(list(
      theme = "7\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "7\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "9\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "9\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    NULL, structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderBottomColour"))


  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,
    NULL, NULL, NULL, NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    structure(list(theme = "6"), .Names = "theme"), structure(list(
      theme = "6"
    ), .Names = "theme"), NULL, NULL, NULL, NULL,
    structure(list(indexed = "64"), .Names = "indexed"), NULL,
    NULL, NULL, NULL, NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "7\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "9\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderTopColour"))



  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, NULL, NULL, NULL, NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), NULL,
    NULL, NULL, NULL, structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), NULL, structure(list(
      indexed = "64"
    ), .Names = "indexed"), NULL, NULL, NULL,
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "3"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), NULL,
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderLeftColour"))



  expected_borders <- list(
    NULL, NULL, NULL, NULL, NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, NULL, NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, NULL, NULL, NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, NULL, NULL, structure(list(theme = "6"), .Names = "theme"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    NULL, structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    NULL, structure(list(theme = "5\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "3"), .Names = "theme"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "6"
    ), .Names = "theme"), NULL, structure(list(
      theme = "6"
    ), .Names = "theme"), structure(list(indexed = "64"), .Names = "indexed"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "7\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "7\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    structure(list(theme = "9\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(theme = "9\" tint=\"-0.249977111117893"), .Names = "theme"),
    structure(list(indexed = "64"), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), structure(list(
      indexed = "64"
    ), .Names = "indexed"), NULL, structure(list(
      theme = "5\" tint=\"-0.249977111117893"
    ), .Names = "theme"),
    NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, structure("9\" tint=\"-0.249977111117893", .Names = "theme"),
    NULL, NULL, NULL, NULL, NULL, NULL
  )

  expect_equal(expected_borders, sapply(styles, "[[", "borderRightColour"))
})
