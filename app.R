library(shiny)
library(rmarkdown)
library(shinyWidgets)
library(dplyr)

# remaking this so that we can style it (more easily) w css later on
fileInputOnlyButton <- function(..., label = "") {
  temp <- fileInput(..., label = label)
  # Cuts away the label
  temp$children[[1]] <- NULL
  # Cut away the input field (after label is cut, this is position 1 now)
  temp$children[[1]]$children[[2]] <- NULL
  # Remove input group classes (makes button flat on one side)
  temp$children[[1]]$attribs$class <- NULL
  temp$children[[1]]$children[[1]]$attribs$class <- NULL
  temp
}

# not nusing fluidpage, since this is more neat
ui <- tabsetPanel(
  id = "panels",
  type = "hidden",
  selected = "main_page",
  tabPanelBody(
    value = "main_page",
    # linking css right into the page
    includeCSS(path = "styles.css"),
      tags$div(
        class = "main_page_content",
          # this is the main heading of the page
          tags$h1(
            tags$strong("ODK2Doc",
              style = "font-size: 100%;
                       padding-bottom: 0px;"
            )
          ),
        tags$div(
          class = "main_text",
          tags$h4(
            HTML(
                        '<br><br>Welcome to <b>ODK2Doc</b>! This is an app to convert ODK/Kobo survey forms to printable, 
                        and easily editable, .docx documents. <br><br> To begin, upload an xls form using 
                        the upload button (<i class="fa fa-upload" style = "color: #8e8d8d;"></i>) below.
                        For best results, make sure your form uses the <b><span class="half_background">
                        <a href = "https://xlsform.org/en/#basic-format/", target="_blank">
                        default</b></span></a> sheet names and a form title that has been added via the 
                        <b><span class="half_background"> <a href = "https://xlsform.org/en/#settings-worksheet", 
                        target="_blank"> settings</b></span></a> sheet. <br><br> To finish, click on the download button (<i class="fa fa-download" style = "color: #8e8d8d;"></i>),
                        and wait for few seconds - your converted form will download as soon as it has been compiled!'
            )
          ),
          style = "font-size: 90%;
                          width: 60%;
                          text-align:justify;
                          line-height: 1.5;"
        ),
        tags$div(
          class = "select_something",
          fileInputOnlyButton(
            inputId     = "file1",
            label       = "Test",
            buttonLabel = list(icon("upload", 
                                    style = "color: #8e8d8d", 
                                    title = "Upload"), 
                               ".xls"),
            accept      = c(".xlsx"),
          ),
        ),
         downloadButton(
           outputId     = "downloadreport",
           icon         = icon("download", 
           style        = "color: #8e8d8d"),
           class        = "select_something",
           style        = HTML("text-decoration: none;"), 
           title        = "Download",
           label        = ".docx",
         ),
        tags$div(
            checkboxInput(inputId = "checkbox",
                          label   = "Add skip logic?", 
                          value   = FALSE, 
                          width   = "100%"),
          ),
        tags$div(
          tags$h3(
          HTML(
            '<b> Note:</b> Still in the testing phase! (This <b>v1.2</b>)<br><br><i class="fa fa-twitter" style = "color: #8e8d8d;"></i>
            <a href = "https://twitter.com/zaeendesouza/", target="_blank">zaeendesouza</span></a>
            <br><i class="fa fa-github" style = "color: #8e8d8d;"></i></i>
            <a href = "https://github.com/zaeendesouza/ODK2Doc", target="_blank">zaeendesouza</span></a>'
          ),
          style          = "font-size: 10px;
                            width: 500px;
                            text-align: center;
                            padding-top: 40px;"
        )
      )
    ),
  )
)


server <- function(input, output) {
  output$downloadreport <-
  downloadHandler(
    filename = "my-odk-form.docx",
      content = function(file) {
        withProgress(message = "Please wait, while we convert your form.", {
          
            if (input$checkbox == TRUE) {
          
            src <- normalizePath("report2.Rmd")
            owd <- setwd(tempdir())
            on.exit(setwd(owd))
            file.copy(from      = src, 
                      to        = "report2.Rmd", 
                      overwrite = T)
            out <- render(input = "report2.Rmd",
                        output_format      = word_document(),
                        params = list(file = input$file1$datapath))
            file.rename(out, file)
            }else{
            src <- normalizePath("report1.Rmd")
            owd <- setwd(tempdir())
            on.exit(setwd(owd))
            file.copy(src, "report1.Rmd", overwrite = T)
            out <- render(input = "report1.Rmd",
                          output_format      = word_document(),
                          params = list(file = input$file1$datapath))
            file.rename(out, file)
          }
        }
      )
    }
  )
}
shiny::shinyApp(ui, server)