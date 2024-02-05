## ----setup, include=FALSE-----------------------------------------------------
knitr::opts_chunk$set(echo = TRUE, comment="")
library(onbrand)
library(officer)
library(magrittr)
library(flextable)

ex_yaml = "rpptx:
  master: Office Theme
  templates:
    title_slide:
      title:
        ph_label:     Title 1
        content_type: text
      sub_title:
        ph_label:     Subtitle 2
        content_type: text"
ex_yaml_tmpfile =  tempfile(fileext=".yaml")
fileConn<-file(ex_yaml_tmpfile)
writeLines(ex_yaml, fileConn)
close(fileConn)

ex_yaml_read = yaml::read_yaml(ex_yaml_tmpfile)

# if_onbrand_hex_png           = system.file(package="onbrand","figures","onbrand_hex.png")
# if_ppt_workflow_output_png   = system.file(package="onbrand","figures","ppt_workflow_output.png")
# if_doc_workflow_output_png   = system.file(package="onbrand","figures","doc_workflow_output.png")
# b64_onbrand_hex_png          = knitrdata::data_encode(if_onbrand_hex_png         , encoding="base64")
# b64_ppt_workflow_output_png  = knitrdata::data_encode(if_ppt_workflow_output_png , encoding="base64")
# b64_doc_workflow_output_png  = knitrdata::data_encode(if_doc_workflow_output_png , encoding="base64")

## -----------------------------------------------------------------------------
obnd = read_template(
       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

## ----echo=FALSE, comment='', message=TRUE, eval=TRUE--------------------------
cat(readLines(ex_yaml_tmpfile) , sep="\n")

## -----------------------------------------------------------------------------
obnd = report_add_slide(obnd,
  template = "title_slide",
  elements = list(
     title     = list( content = "Onbrand PowerPoint Example",
                       type    = "text"),
     sub_title = list( content = "Workflow Abstraction",
                       type    = "text")))

## -----------------------------------------------------------------------------
bl = c("1", "This is first level bullet",
       "2", "sub-bullet",
       "3", "sub-sub-bullet",
       "3", "can't have just one sub-sub-bullet",
       "2", "same goes for sub-bullets",
       "1", "Another first level bullet")

## -----------------------------------------------------------------------------
obnd = report_add_slide(obnd,
  template = "content_list",
  elements = list(
     title        = list( content = "Adding List Content",
                          type    = "text"),
     content_body = list( content = bl,
                          type    = "list")))

## -----------------------------------------------------------------------------
library(ggplot2)
p = ggplot() + annotate("text", x=0, y=0, label = "picture example")
imgfile = tempfile(pattern="image", fileext=".png")
ggsave(filename=imgfile, plot=p, height=5.15, width=9, units="in")

## -----------------------------------------------------------------------------
obnd = report_add_slide(obnd,
  template = "two_content_header_text",
  elements = list(
     title                = list(content  = "Adding Images Content",
                                 type     = "text"),
     content_left_header  = list(content  ="ggplot object",
                                 type     = "text"),
     content_left         = list(content  = p,
                                 type     = "ggplot"),
     content_right_header = list(content  ="image file",
                                 type     = "text"),
     content_right        = list(content  = imgfile,
                                 type     = "imagefile")))

## -----------------------------------------------------------------------------
tdf = data.frame(Parameters = c("Length", "Width", "Height"),
                 Values     = 1:3,
                 Units      = c("m", "m", "m") )

## -----------------------------------------------------------------------------
tab_cont = list(table = tdf)
obnd = report_add_slide(obnd,
  template = "content_text", 
  elements = list(
     title         = list( content      = "Tables",
                           type         = "text"),
     sub_title     = list( content      = "Creating PowerPoint Table",
                           type         = "text"),
     content_body  = list( content      = tab_cont,
                           type         = "table")))

## -----------------------------------------------------------------------------
tab_ft = list(table         = tdf,
              header_format = "md",
              header_top    = list(Parameters = "Name^2^",
                                   Values     = "*Value*",
                                   Units      = "**Units**"),
              cwidth        = 0.8,
              table_autofit = TRUE,
              table_theme   = "theme_zebra")

## -----------------------------------------------------------------------------
tab_fto = flextable(tdf)                      

## -----------------------------------------------------------------------------
obnd = report_add_slide(obnd,
  template = "two_content_header_text",
  elements = list(
     title                 = list( content      = "Tables",
                                   type         = "text"),
     sub_title             = list( content      = "flextables can be created in two ways",
                                   type         = "text"),
     content_left_header   = list( content      = 'using "flextable"',
                                   type         = "text"),
     content_left          = list( content      = tab_ft,
                                   type         = "flextable"),
     content_right_header  = list( content      = 'using "flextable_objecct"',
                                   type         = "text"),
     content_right         = list( content      = tab_fto,
                                   type         = "flextable_object")))

## -----------------------------------------------------------------------------
obnd = report_add_slide(obnd,
  template = "two_content_header_list",
  elements = list(
     title         = list( content      = "Combining elements, user defined locations, and formatting",
                           type         = "text")),
  user_location = list(
     txt_example   =
       list( content        = fpar(ftext("This is formatted text", fp_text(color="green", font.size=24))),
             type           = "text",
             start          = c(0,  .25),
             stop           = c(.25,.35)),
     large_figure  =
       list( content        = p,
             type           = "ggplot",
             start          = c(.25,.25),
             stop           = c(.99,.99)),
     flextable_obj  =
       list( content        = tab_fto,
             type           = "flextable_object",
             start          = c(0,.75),
             stop           = c(.25,.95)),
     small_figure  =
       list( content        = p,
             type           = "ggplot",
             start          = c(0,  .35),
             stop           = c(.25,.74))
  )
)

## ----eval=TRUE, echo=FALSE, results='hide', message=FALSE---------------------
save_report(obnd,  file.path(tempdir(), "vignette_presentation.pptx"))

## ----message=TRUE-------------------------------------------------------------
details = template_details(obnd) 

## -----------------------------------------------------------------------------
ph = fph(obnd, "two_content_header_text", "content_left_header")$pl

## -----------------------------------------------------------------------------
obnd = read_template(
  template = file.path(system.file(package="onbrand"), "templates", "report.docx"),
  mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "toc",
  content  = list(level=3))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "toc",
  content  = list(style="Table_Caption"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "toc",
  content  = list(style="Figure_Caption"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text="Text with no style specified will use the doc_def text format. This is a 'ph' placehoder for text: ===BODY-TEXT-EXAMPLE=== [see Placeholder text section below]"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text  ="First level header",
                  style = "Heading_1"))
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text  ="Second level header",
                  style = "Heading_2"))
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text  ="Third level header",
                  style = "Heading_3"))

## -----------------------------------------------------------------------------
library(officer)

fpartext = fpar(
ftext("Formatted text can be created using the ", prop=NULL),
ftext("fpar ", prop=fp_text(color="green")),
ftext("command from the officer package.", prop=NULL))

obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text   = fpartext, 
                  format = "fpar",
                  style  = "Normal"))

mdtext = "Formatted text can be created using **<color:green>markdown</color>** formatting"
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text   = mdtext,
                  format = "md",
                  style  = "Normal"))


## -----------------------------------------------------------------------------
p = ggplot() + annotate("text", x=0, y=0, label = "picture example")
imgfile = tempfile(pattern="image", fileext=".png")
ggsave(filename=imgfile, plot=p, height=5.15, width=9, units="in")

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "imagefile",
  content  = list(image   = imgfile,
                  caption = "This is an example of an image from a file."))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "ggplot",
  content  = list(image   = p,
                  caption = "This is an example of an image from a ggplot object."))

## -----------------------------------------------------------------------------
tdf =    data.frame(Parameters = c("Length", "Width", "Height"),
                    Values     = 1:3,
                    Units      = c("m", "m", "m") )

## -----------------------------------------------------------------------------
tab_cont = list(table   = tdf,
                caption = "Word Table.")
obnd = report_add_doc_content(obnd,
  type     = "table",
  content  = tab_cont)

## -----------------------------------------------------------------------------
tab_ft = list(table         = tdf,
              header_format = "md",
              header_top    = list(Parameters = "Name^2^",
                                   Values     = "*Value*",
                                   Units      = "**Units**"),
              cwidth        = 0.8,
              table_autofit = TRUE,
              caption       = "Flextable from onbrand abstraction",
              table_theme   = "theme_zebra")
obnd = report_add_doc_content(obnd,
  type     = "flextable",
  content  = tab_ft)   

## -----------------------------------------------------------------------------
tab_fto = flextable(tdf)                      
obnd = report_add_doc_content(obnd,
  type     = "flextable_object",
  content  = list(ft=tab_fto,
                  caption  = "Flextable object created by the user."))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
     type     = "section",
     content  = list(section_type  ="portrait"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "ggplot",
  content  = list(image   = p,
                  height  = 2.5,
                  width   = 9,
                  caption = "This is a landscape figure."))

## -----------------------------------------------------------------------------

obnd = report_add_doc_content(obnd,
  type     = "section",
  content  = list(section_type  ="landscape",
                  height        = 8,
                  width         = 10))
 

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "text",
  content  = list(text    = paste(rep("Some two column text.", 200), collapse=" ")))


## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "section",
  content  = list(section_type  ="columns",
                  widths        = c(3,3)))


## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type          = "text",
  content       = list(text    = "Back to regular portrait"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "ph",
  content  = list(name     = "BODY-TEXT-EXAMPLE",
                  value    = "Swaps the placeholder with this Text",
                  location = "body"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "ph",
  content  = list(name     = "FOOTERLEFT",
                  value    = "Text Swapped with Footer Placeholder",
                  location = "footer"))
obnd = report_add_doc_content(obnd,
  type     = "ph",
  content  = list(name     = "HEADERLEFT",
                  value    = "Text Swapped with Header Placeholder",
                  location = "header"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "section",
  content  = list(section_type  ="portrait"))

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "ggplot",
  content  = list(image           = p,
                  notes_format    = "text",
                  key             = "ex_fig_text",
                  notes           = "This figure shows how to use text captions _and_ notes",
                  caption_format  = "text",
                  caption         = "Multi-page figure (page 1)"))

obnd = report_add_doc_content(obnd,
  type     = "break",
  content  = NULL)

obnd = report_add_doc_content(obnd,
  type     = "ggplot",
  content  = list(image           = p,
                  notes_format    = "text",
                  key             = "ex_fig_text",
                  notes           = "This figure shows how to use text captions _and_ notes",
                  caption_format  = "text",
                  caption         = "Multi-page figure (page 2)"))

## ----eval=FALSE, echo=FALSE---------------------------------------------------
#  of = tempfile(fileext=".docx")
#  save_report(obnd, of)

## ----message=TRUE-------------------------------------------------------------
details = template_details(obnd) 

## -----------------------------------------------------------------------------
st = fst(obnd, "Heading_3")
#Word style name
wsn = st$wsn
# Default font format
dff = st$dff

## -----------------------------------------------------------------------------
data = data.frame(property = c("mean",   "variance"),
                  length     = c(200,      0.13),
                  width      = c(12,       0.05),
                  area       = c(240,      0.11),
                  volume     = c(1200,     0.32))


## -----------------------------------------------------------------------------
header = list(property = c("",             ""),
              length   = c("Length",       "cm"),
              width    = c("Wdith",        "cm"),
              area     = c("Area",         "cm2"),
              volume   = c("Volume",       "cm3"))

ft = flextable::flextable(data)                     %>% 
     flextable::delete_part(part = "header")        %>%
     flextable::add_header(values =as.list(header)) %>%
     flextable::theme_zebra()

## ----echo=FALSE---------------------------------------------------------------
htmltools_value(ft)

## -----------------------------------------------------------------------------
dft      = fetch_md_def(obnd, style="Table_Labels")$md_def
dft_body = fetch_md_def(obnd, style="Table")$md_def

## -----------------------------------------------------------------------------
ft = ft %>%
  flextable::compose(j     = "area",
          part  = "header", 
          value = c(md_to_oo("Area", dft)$oo, md_to_oo("cm^2^", dft)$oo))   %>%
  flextable::compose(j     = "volume", 
          part  = "header",
          value = c(md_to_oo("Volume", dft)$oo, md_to_oo("cm^3^", dft)$oo)) %>%
  flextable::compose(j     = "property", 
          i     = match("mean", data$property),                        
          part  = "body",  
          value = c(md_to_oo("**<ff:symbol>m</ff>**", dft_body)$oo))    %>%
  flextable::compose(j     = "property",
          i     = match("variance", data$property), 
          part  = "body",                                                     
          value = c(md_to_oo("**<ff:symbol>s</ff>**^**2**^", dft_body)$oo))

## ----echo=FALSE---------------------------------------------------------------
htmltools_value(ft)

## -----------------------------------------------------------------------------
obnd = report_add_doc_content(obnd,
  type     = "flextable_object",
  content  = list(ft=tab_fto,
                  caption  = "Flextable object with custom Markdown - created by the user."))

## ----echo=FALSE---------------------------------------------------------------
rpt = fetch_officer_object(obnd)$rpt

## ----echo=FALSE---------------------------------------------------------------
obnd = set_officer_object(obnd, rpt)

## ----echo=FALSE, comment='', message=TRUE, eval=TRUE--------------------------
cat(readLines(file.path(system.file(package="onbrand"), "templates", "report.yaml")) , sep="\n")

