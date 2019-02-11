library(shiny)
library(openxlsx)
library(plyr)
library(dplyr)
library(rclipboard)
library(shinyjs)
library(xtable)
library(knitr)
library(kableExtra)

#Setting the CSS for the App
appCSS <- "#welcome {
text-align: center;
}

#copyButton {
font-size: 20px;
background-color: orange;
}

#download_sps {
font-size: 15px;
background-color: SkyBlue;
}

#syntax_table th, td{
border: 1px white solid;
font-size: 16px;
}

#no_upload {
color: DodgerBlue;
text-align: center;
}"

### UI configuration ###
ui <- fluidPage(
  
  #Activate necessary packages
  rclipboardSetup(),
  useShinyjs(),
  inlineCSS(appCSS),
  
  # Prepare loaing screen
  div(
    id = "loading-content",
    h2("Loading...")
  ),
  
  hidden(
    div(
      id = "app-content",
      
      # Main code of UI
      titlePanel(title = "SPSS Syntax Generator"),
      
      sidebarLayout(
        sidebarPanel(
          
          textOutput("test_string"),
          
          #Upload Button
          fileInput("uploaded_file", "Browse to your .xslx file", accept = ".xlsx"),
          
          #Get SPSS Syntax Button
          actionButton("syntax", "Get SPSS Syntax"),
          
          p(),
          
          #Example File
          downloadLink("example_file", "Download an example file"),
          
          p(),
          
          #Link to Tutorial
          a("Tutorial", href="https://www.youtube.com/watch?v=a61Fr7DdgaQ&t=1s", target = "_blank")
        ),
        
        mainPanel(
          
          # Show Logo
          tags$img(id = "logo", src="syntax_generator_logo_crop.png", width="20%", height="20%",
                   style="display: block; margin-left: auto; margin-right: auto;"),
          
          #Warning Message if Get SPSS Syntax button hit without a file uploaded
          h2(textOutput("no_upload")),
          
          #Welcome message
          h4(id = "welcome", "Here you can generate your SPSS Syntax in a couple of seconds.
             If you don't know what this is about, you can find a link to the tutorial on the left"),
          

          #Copy to clipboard Button
          uiOutput("copy_button"),
          
          p(),
          
          #Download .sps File Button
          downloadButton("download_sps", "Download .sps file"),
          
          #Show the Syntax Table
          tableOutput("syntax_table")
          )
        )
    )
  ))

### Server configuration ###
server <- function(input, output) {
  
  #Simulate work being done for 1 second for loading screen
  Sys.sleep(1)
  
  #Hide the loading message when the rest of the server function has executed
  hide(id = "loading-content")
  show("app-content")
  
  #Hide the download .sps Button for now
  hide("download_sps")
  
  #Set option for kable to replace NAs
  options(knitr.kable.NA = "")
  
  
  
  ### Code to  be executed when Get SPSS Syntax Button gets hit ###
  observeEvent(input$syntax,  {
    
    
    #Check if an uploaded file is present
    if (is.null(input$uploaded_file) || substr(input$uploaded_file$name, 
                                               nchar(input$uploaded_file$name) - 4,
                                               nchar(input$uploaded_file$name)) != ".xlsx") {
      output$no_upload <- renderText("Please upload an Excel file")
    } else {
      
      #Convert uploaded file to data.frames
      variables_table <- read.xlsx(input$uploaded_file$datapath, 1)
      
      values_table <- read.xlsx(input$uploaded_file$datapath, 2)
      
      #Put together the actual Syntax for the Variables if data.frame is present
      if (is.na(variables_table[1,1]) == FALSE) {
        
        #Rename Variable Names
        syntax_var_table <- data.frame(V1 = "Rename Variable",  stringsAsFactors = FALSE)
        
        for (i in 1:nrow(variables_table)) {
          syntax_var_table[i+1,1] <- paste(variables_table[i,1], "=", variables_table[i,2])
        }
        
        syntax_var_table <- rbind(syntax_var_table, rbind(".", NA))
        
        #Label Variables
        syntax_var_val_table <- data.frame(V1 = "Variable Label",  stringsAsFactors = FALSE)
        
        for (i in 1:nrow(variables_table)) {
          syntax_var_val_table[i+1,1] <- paste0(variables_table[i,2], ' "', variables_table[i,3], '"')
        }
        
        syntax_var_val_table <- rbind(syntax_var_val_table, rbind(".", NA))
        
        #Bind variables and Variables Value Table
        syntax_final_table <- rbind(syntax_var_table, syntax_var_val_table)
      } else {
        
        #else, make an empty data.frame
        syntax_final_table <- data.frame()
      }
      
      #Values labeling
      #Put together the actual Syntax for the Values if data.frame is present
      if (is.na(values_table[1,1]) == FALSE) {
        
        syntax_val_table <- data.frame(V1 = "Value Labels",  stringsAsFactors = FALSE)
        
        for (i in 1:nrow(values_table)) {
          
          syntax_val_table[i+1,1] <- values_table[i,1]
          syntax_val_table[i+1,2] <- values_table[i,2]
          syntax_val_table[i+1,3] <- paste0('"', values_table[i,3], '"')
          
        }
        
        for (i in 1:nrow(syntax_val_table)) {
          if (is.na(syntax_val_table[i,1]) == TRUE & !is.na(syntax_val_table[i+1,1])) {
            syntax_val_table[i,4] <- " /"
          }
        }
        
        syntax_val_table[nrow(syntax_val_table) + 1, 1] <- "."
        
        #Binding everything together
        syntax_final_table <- rbind.fill(syntax_final_table, syntax_val_table)
        
      }
      
      #Prepare output of Syntax Table
      output$syntax_table <- function() {
        kable(syntax_final_table, "html", row.names = FALSE, col.names = c(rep("", ncol(syntax_final_table)))) %>%
          kable_styling(full_width = TRUE)
      }
      
      ### Copy to clipboard Button###
      output$copy_button <- renderUI({
        tc <- textConnection("table_to_copy", open = "w", local = FALSE, encoding = "UTF-8") #open connection
        sink(tc)   #divert output to tc connection
        syntax_final_table <- syntax_final_table %>% mutate_all(as.character) #turn all columns to character to correctly replace NAs
        print(syntax_final_table, na.print = "")  #print in str string instead of console
        sink()     #set the output back to console
        close(tc)  #close connection
        table_to_copy <- substr(table_to_copy[-1], floor(log10(length(table_to_copy))) + 3, nchar(table_to_copy[1])) #get rid of the row numbers that come with print and the column names
        table_to_copy <- paste0(table_to_copy,collapse="\n") #we build a proper unique string with pipes and new lines
        rclipButton("copyButton", "Copy to Clipboard", table_to_copy , icon("clipboard")) #actual ouput for the Button
      })
      
      ### Add sps download button ###
      
      #unhide the button
      show("download_sps")
      
      output$download_sps <- downloadHandler(
        filename = function() {
          paste("syntax_", Sys.Date(), ".sps", sep="")
        },
        content = function(file) {
          write.table(syntax_final_table, sep="\t", row.names=FALSE, col.names = FALSE, na = "",
                      quote = FALSE, fileEncoding="UTF-8", file)
        }
      )
      
      #Content to be hidden when everything goes as planned
      hide("no_upload")
      hide("welcome")
      hide("logo")
    }
  }
  )
  
  ### Download example file ####
  
  #Output for example excel file
  output$example_file <- downloadHandler(
    filename <- function() {
      paste("example_file", "xlsx", sep = ".")
    },
    content <- function(file) {
      file.copy("example_file.xlsx", file)
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )
}

### Knit everything together ###
shinyApp(ui = ui, server = server)
