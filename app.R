library(shiny)
library(rmarkdown)
library(shinycssloaders)
library(dplyr)
library(waiter)

# remaking this so that we can style it (more easily) w css later on
fileInputOnlyButton <- function(..., label = "") {
  temp <- fileInput(..., label = label)
  # Cut away the label
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
      class = "main_page_container",
      tags$div(
        class = "main_page_content",
        tags$div(
          class = "some_text",

          # this is the main heading of the page
          tags$h1(
            tags$strong("ODK2Doc",
              style = "font-size: 100%"
            )
          )
        ),
        tags$div(
          class = "main_text",
          tags$h4(
            HTML(
              'Load an xls form using the upload button
                        (<i class="fa fa-upload" style = "color: #8e8d8d;"></i>).
                        Make sure your form uses the <b><span class="half_background">
                        <a href = "https://xlsform.org/en/#basic-format/", target="_blank">
                        default</b></span></a> sheet names. You should also add a form title via the <b><span class="half_background">
                        <a href = "https://xlsform.org/en/#settings-worksheet", target="_blank">
                        settings</b></span></a> sheet. <br><br> Once you have uploaded a form,
                        click the download button (<i class="fa fa-download" style = "color: #8e8d8d;"></i>),
                        and wait for few seconds; your converted form will download
                        as soon as the conversion is over!'
            )
          ),
          style = "font-size: 100%;
                          width: 50%;
                          text-align:justify;
                          line-height: 1.7"
        ),
        tags$div(
          class = "select_something",
          fileInputOnlyButton(
            inputId     = "file1",
            label       = "Test",
            buttonLabel = list(icon("upload", style = "color: #8e8d8d"), ".xls"),
            accept      = c(".xlsx"),
          ),
        ),
        downloadButton(
          outputId = "downloadreport",
          icon     = icon("download", style = "color: #8e8d8d"),
          class    = "select_something",
          style    = HTML("text-decoration: none;"),
          label    = ".doc",
        ),
        tags$h3(
          HTML(
            '<b> Note:</b> Still in testing phase! (This <b>v.1</b>).<br>The next version will contain more features! <br><br><i class="fa fa-twitter" style = "color: #8e8d8d;"></i>
                     <a href = "https://twitter.com/zaeendesouza/", target="_blank">zaeendesouza</span></a>
                     <br><i class="fa fa-github" style = "color: #8e8d8d;"></i></i>
                     <a href = "https://github.com/zaeendesouza", target="_blank">zaeendesouza</span></a>'
          ),
          style = "font-size: 10px;
                         width: 900px;
                         text-align: center;
                        padding-top: 60px;"
        )
      )
    ),
  ),
)


server <- function(input, output) {
  output$downloadreport <-
    
    downloadHandler(
      
      filename = "my-odk-form.docx",
      
      content = function(file) {
        
        withProgress(message = "Please wait, while we convert your form.", {
          
          src <- normalizePath("report.Rmd")
          
          owd <- setwd(tempdir())
          
          on.exit(setwd(owd))
          
          file.copy(src, "report.Rmd", overwrite = TRUE)
          
          out <- render(
            "report.Rmd",
            output_format      = word_document(),
            params = list(file = input$file1$datapath)
          )
          
          file.rename(out, file)
          
        }
      )
    }
  )
}
shiny::shinyApp(ui, server)