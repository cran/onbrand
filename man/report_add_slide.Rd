% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/report_add_slide.R
\name{report_add_slide}
\alias{report_add_slide}
\title{Add Slide and Content}
\usage{
report_add_slide(obnd, template = NULL, elements = NULL, verbose = TRUE)
}
\arguments{
\item{obnd}{onbrand report object}

\item{template}{Name of slide template to use (name from templates in yaml mapping file)}

\item{elements}{Content and type for each placeholder you wish to fill for
this slide: This is a list with names set to palceholders for the specified
tempalte. Each placeholder is a list and should have a content element and a
type element (see Details below).}

\item{verbose}{Boolean variable when set to TRUE (default) messages will be
displayed on the terminal; Messages will be included in the returned onbrand
object.}
}
\value{
onbrand report object with either the content added or isgood set
to FALSE with any messages in the msgs field.
}
\description{
Creates a report slide and populates the content
}
\details{
For example consider the mapping information for the slide
template \code{title_slide} with the two place holders \code{title} and
\code{subtitle}.

\preformatted{
rpptx:
 master: Office Theme
 templates:
   title_slide:
     title:
       type:         ctrTitle
       index:        1
       ph_label:     Title 1
       content_type: text
     subtitle:
       type:         subTitle
       index:        1
       ph_label:     Subtitle 2
       content_type: text
}

This shows how to populate a title slide with text:

\preformatted{
obnd = report_add_slide(obnd,
 template = "title_slide",
 elements = list(
    title     = list( content      = "Slide Title",
                      type         = "text"),
    subtitle  = list( content      = "Subtitle",
                      type         = "text")))
}

See the function \code{\link{add_pptx_ph_content}} for a list of
allowed values for \code{type}. Note that if mapping defines the
\code{content_type} as \code{text}, you cannot use a \code{list} type.
Similarly, if the \code{content_type} is defined as \code{list}, you
cannot use a \code{text} type.
}
\examples{
obnd = read_template(
      template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
      mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))

obnd = report_add_slide(obnd,
 template = "content_text",
 elements = list(
    title         = list( content      = "Text Example",
                          type         = "text"),
    sub_title     = list( content      = "Adding a slide with a block of text",
                          type         = "text"),
    content_body  = list( content      = "A block of text",
                          type         = "text")))

}
\seealso{
\code{\link{add_pptx_ph_content}}
}
