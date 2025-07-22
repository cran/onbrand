## ----setup, include=FALSE-----------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


# Colors used in graphics to describe different elements
c_orange = "ffa500ff"
c_green  = "44aa00ff"
c_blue   = "0000ffff"
c_purple = "7137c8ff"

ex_yaml = "tree colors:
  roots:
    - white 
    - brown
  trunk:
    bark: brown
  crown:
    branches:
      leaves:  green
      flowers: red"
ex_yaml_tmpfile =  tempfile(fileext=".yaml")
fileConn<-file(ex_yaml_tmpfile)
writeLines(ex_yaml, fileConn)
close(fileConn)

ex_yaml_read = yaml::read_yaml(ex_yaml_tmpfile)

# if_onbrand_hex_png                  = system.file(package="onbrand","figures","onbrand_hex.png")
# if_example_layout_ppt_master_png    = system.file(package="onbrand","figures","example_layout_ppt_master.png")
# if_example_layout_yaml_pptx_png     = system.file(package="onbrand","figures","example_layout_yaml_pptx.png")
# if_example_layout_docx_png          = system.file(package="onbrand","figures","example_layout_docx.png")
# b64_onbrand_hex_png                 = knitrdata::data_encode(if_onbrand_hex_png               , encoding="base64")
# b64_example_layout_ppt_master_png   = knitrdata::data_encode(if_example_layout_ppt_master_png , encoding="base64")
# b64_example_layout_yaml_pptx_png    = knitrdata::data_encode(if_example_layout_yaml_pptx_png  , encoding="base64")
# b64_example_layout_docx_png         = knitrdata::data_encode(if_example_layout_docx_png       , encoding="base64")

## ----eval=FALSE---------------------------------------------------------------
# library(onbrand)
# file.copy(system.file(package="onbrand","examples","example.pptx"), ".", overwrite = TRUE)
# file.copy(system.file(package="onbrand","examples","example.docx"), ".", overwrite = TRUE)
# file.copy(system.file(package="onbrand","examples","example.yaml"), ".", overwrite = TRUE)

## ----echo=FALSE, comment='', message=TRUE, eval=TRUE--------------------------
cat(readLines(ex_yaml_tmpfile) , sep="\n")

## ----eval=FALSE---------------------------------------------------------------
# library(onbrand)
# vlres = view_layout(template    = "example.pptx",
#                     output_file = "example_layout.pptx")

## ----eval=FALSE---------------------------------------------------------------
#  obnd = read_template(template = "example.pptx",
#                       mapping  = "example.yaml")

## ----eval=FALSE---------------------------------------------------------------
# obnd = preview_template(obnd)
# save_report(obnd, "example_preview.pptx")

## ----eval=FALSE---------------------------------------------------------------
# library(onbrand)
# vlres = view_layout(template    = "example.docx",
#                     output_file = "example_layout.docx")

## ----eval=FALSE---------------------------------------------------------------
# obnd = read_template(template = "example.docx",
#                      mapping  = "example.yaml")

## ----eval=FALSE---------------------------------------------------------------
# obnd = preview_template(obnd)
# save_report(obnd, "example_preview.docx")

## ----echo=FALSE, comment='', message=TRUE, eval=TRUE--------------------------
cat(readLines(file.path(system.file(package="onbrand"), "examples", "example.yaml")) , sep="\n")

