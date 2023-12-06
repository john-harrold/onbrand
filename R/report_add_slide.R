#'@export
#'@title Add Slide and Content
#'@description Creates a report slide and populates the content in
#'placeholders and arbitrary locations.
#'
#'@param obnd onbrand report object
#'@param template Name of slide template to use (name from templates in yaml mapping file)
#'@param elements Content and type for each placeholder you wish to fill for
#'this slide: This is a list with names set to placeholders for the specified
#'template. Each placeholder is a list and should have a content element and a
#'type element (see Details below).
#'@param user_location List with arbitrary element names (see Details below)
#'@param verbose Boolean variable when set to TRUE (default) messages will be
#'displayed on the terminal; Messages will be included in the returned onbrand
#'object.
#'
#'@return onbrand report object with either the content added or isgood set
#'to FALSE with any messages in the msgs field. The isgood value is a Boolean variable
#'indicating the current state of the object.
#'
#'@details
#'To add content based on placeholder names consider the mapping information for the slide
#'template \code{title_slide} with the two place holders \code{title} and
#'\code{subtitle}.
#'
#'\preformatted{rpptx:
#'  master: Office Theme
#'  templates:
#'    title_slide:
#'      title:
#'        type:         ctrTitle
#'        index:        1
#'        ph_label:     Title 1
#'        content_type: text
#'      subtitle:
#'        type:         subTitle
#'        index:        1
#'        ph_label:     Subtitle 2
#'        content_type: text}
#'
#'This shows how to populate a title slide with text:
#'
#'\preformatted{obnd = report_add_slide(obnd,
#'  template = "title_slide",
#'  elements = list(
#'     title     = list( content      = "Slide Title",
#'                       type         = "text"),
#'     subtitle  = list( content      = "Subtitle",
#'                       type         = "text")))}
#'
#'
#'To add content based on user defined locations you need to supply a list
#'with the content, type, starting point and stopping point. You can use any
#'template you wish, and you need to populate the \code{user_location} input.
#'This consists of lists. The name of these lists can be arbitrary
#'(text_example and fig_example below). Each list has a content and type, this
#'is the same used in elements above. The start and stop each represent x and
#'y coordinates. This is the fraction of the width and height of the slide
#'measured from the upper left. So the \code{start = c(0.5, 0)} below means
#'the box holding that content would start at the middle of the slide width
#'and the top of the slide.
#'
#'\preformatted{ #'obnd = report_add_slide(obnd,
#'  template = "two_content_header_text",
#'  user_location = list(
#'     text_example  = list( content      = "This is text",
#'                           type         = "text",
#'                           start        = c(.01,.02),
#'                           stop         = c(.3,.15)),
#'     fig_example   = list( content      = ggplot2::ggplot(),
#'                           type         = "ggplot",
#'                           start        = c(.5,0),
#'                           stop         = c(1,.5))
#'  )
#')}
#'
#'See the function \code{\link{add_pptx_ph_content}} for a list of
#'allowed values for \code{type}. Note that if mapping defines the
#'\code{content_type} as \code{text}, you cannot use a \code{list} type.
#'Similarly, if the \code{content_type} is defined as \code{list}, you
#'cannot use a \code{text} type.
#'
#'@seealso \code{\link{add_pptx_ph_content}} \code{\link{view_layout}}
#'
#'@examples
#'obnd = read_template(
#'       template = file.path(system.file(package="onbrand"), "templates", "report.pptx"),
#'       mapping  = file.path(system.file(package="onbrand"), "templates", "report.yaml"))
#'
#'# Adding content based on placeholder elements
#'obnd = report_add_slide(obnd,
#'  template = "content_text",
#'  elements = list(
#'     title         = list( content      = "Text Example",
#'                           type         = "text"),
#'     sub_title     = list( content      = "Adding a slide with a block of text",
#'                           type         = "text"),
#'     content_body  = list( content      = "A block of text",
#'                           type         = "text")))
#'
#'# Adding content based on specified locations
#'obnd = report_add_slide(obnd,
#'  template = "two_content_header_text",
#'  user_location = list(
#'     text_example  = list( content      = "This is text",
#'                           type         = "text",
#'                           start        = c(.01,.02),
#'                           stop         = c(.3,.15))))
report_add_slide = function (obnd,
                             template      = NULL,
                             elements      = NULL,
                             user_location = NULL,
                             verbose       = TRUE){
  msgs          = c()
  isgood        = TRUE
  general_types = c("imagefile", "ggplot", "table", "flextable", "flextable_object")

  #-----
  # This is a set of basic tests on user input
  if(obnd[["isgood"]]){
    if(obnd[["rpttype"]] != "PowerPoint"){
      isgood = FALSE
      msgs = c(msgs, paste0("The povided onbrand object is for rpttype >", obnd[["rpttype"]], "<"))
      msgs = c(msgs, "and should be PowerPoint")
    }
  } else {
    isgood = FALSE
    msgs = c(msgs, "Bad onbrand object supplied")
  }
  if(is.null(template)){
    isgood = FALSE
    msgs = c(msgs, "No slide template was provided")
  }
  if(is.null(elements) & is.null(user_location)){
    isgood = FALSE
    msgs = c(msgs, "No slide content in elements or user_location were provided")
  }
  #-----
  # If we pass the basic tests then we start walking through the user supplied
  # information to make sure it makes sense in the context of the current
  # templates
  td = NULL
  if(isgood){
    if(template %in% names(obnd[["meta"]][["rpptx"]][["templates"]])){
      # this holds the template details for the current layout:
      td = obnd[["meta"]][["rpptx"]][["templates"]][[template]]
      # If the template exists we test each placeholder next
      if(all(names(elements) %in% names(td))){
        # now we check the elements that are specified as text/list
        for(phname in names(elements)){
          # Making sure the content types are correct
          allowed_types = c(td[[phname]][["content_type"]], general_types)
          if(!(elements[[phname]][["type"]] %in% allowed_types)){
            isgood = FALSE
            msgs = c(msgs, paste0("The content type >", elements[[phname]][["type"]], "< is not allowed with"))
            msgs = c(msgs, paste0("the layout template >", template, "<"))
            msgs = c(msgs, paste0("The allowed values of type are:"))
            msgs = c(msgs, paste0("  ", paste(allowed_types, collapse=", ")))
          }
        }
      } else {
        isgood = FALSE
        msgs = c(msgs, paste0("The following placeholder elements were specified but"))
        msgs = c(msgs, paste0("not found in the mapping file: "))
        msgs = c(msgs, paste0("  ",
          paste(names(elements)[!(names(elements) %in% names(td))], collapse =", ")))
      }

      # Checking content added by location
      for(locname in names(user_location)){
        if( length(user_location[[locname]][["start"]]) == 2 &
            is.numeric(user_location[[locname]][["start"]])  &
            length(user_location[[locname]][["start"]]) == 2 &
            is.numeric(user_location[[locname]][["start"]])){
          if(user_location[[locname]][["start"]][1] > user_location[[locname]][["stop"]][1]){
            isgood = FALSE
            msgs = c(msgs, paste0("Problem with user location >", locname, "<"))
            msgs = c(msgs, paste0("x start is > x stop"))
          }
          if(user_location[[locname]][["start"]][2] > user_location[[locname]][["stop"]][2]){
            isgood = FALSE
            msgs = c(msgs, paste0("Problem with user location >", locname, "<"))
            msgs = c(msgs, paste0("y start is > y stop"))
          }
        } else{
          isgood = FALSE
          msgs = c(msgs, paste0("Problem with user location >", locname, "< the start"))
          msgs = c(msgs, paste0("and stop values must be numeric of length two."))
        }
      }
    }  else {
      isgood = FALSE
      msgs = c(msgs, paste0("The slide template >", template, "< was not found"))
      msgs = c(msgs, paste0("in the mapping file."))
    }

  }


  #-------
  # If all that above checks out then we add the reporting elements
  if(isgood){
    # First we add the new slide according to the specified template
    obnd[["rpt"]] = officer::add_slide(x      = obnd[["rpt"]],
                       layout = template,
                       master = obnd[["meta"]][["rpptx"]][["master"]])

    # Now we walk through each specified placeholder and add it
    for(phname in names(elements)){
      if(is.null(td[[phname]][["ph_label"]])){
        msgs = c(msgs, paste0("The named placeholder >", phname,"< is defined in the onbrand mapping file but is NULL"))
        msgs = c(msgs, "This can happen when a reporting workflow has a placeholder but it is not implemented")
        msgs = c(msgs, "a specific template. This element will be skipped and not added to the report.")
      } else {
        obnd = add_pptx_ph_content(obnd = obnd,
                                   content      = elements[[phname]][["content"]],
                                   content_type = elements[[phname]][["type"]],
                                   ph_label     = td[[phname]][["ph_label"]],
                                   verbose      = verbose)
      }
    }
    # Same thing here but we add content by location
    if(!is.null(names(user_location))){
      # We need these dims to calculate the offsets below
      slide_dims = officer::slide_size(obnd[["rpt"]])
      sw  = slide_dims[["width"]]
      sh  = slide_dims[["height"]]

      for(locname in names(user_location)){
        xstart =  user_location[[locname]][["start"]][1]
        ystart =  user_location[[locname]][["start"]][2]
        xstop  =  user_location[[locname]][["stop"]][1]
        ystop  =  user_location[[locname]][["stop"]][2]

        left   = xstart*sw
        top    = ystart*sh
        width  = (xstop - xstart)*sw
        height = (ystop - ystart)*sh

        newlabel = digest::digest(user_location[[locname]], algo=c("md5"))

        tmp_location = officer::ph_location(
          left     = left,
          top      = top,
          width    = width,
          height   = height,
          newlabel = newlabel)

        obnd = add_pptx_ph_content(obnd = obnd,
          content       = user_location[[locname]][["content"]],
          content_type  = user_location[[locname]][["type"]],
          user_location = tmp_location,
          verbose       = verbose)
      }
    }
  }
  #-------
  # If errors were encountered we make sure that the state of the reporting
  # object is set to false
  if(!isgood){
    obnd[["isgood"]] = FALSE
    msgs = c(msgs, paste0("mapping file: ", obnd[["mapping"]]))
    msgs = c(msgs, "onbrand::report_add_slide()")
  }

  # Dumping the messages if verbose is turned on:
  if(verbose & !is.null(msgs)){
    message(paste(msgs, collapse="\n"))
  }

  if(!isgood){
    stop("Unable to add PowerPoint Slide. See the messages above for details.")
  }

  # Adding any messages to the report object
  obnd[["msgs"]] = msgs


obnd}
