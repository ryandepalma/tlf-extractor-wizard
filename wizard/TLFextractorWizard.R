################################################################################                     
#                     ==============================                           #      
#                         TLF EXTRACTOR WIZARD                                 #     /\  
#                        ----------------------                                #    /__\ 
# Description: This shiny app extracts all Tables, Listings, and Figures       #    {00} *
#              from a SAP Shell and generates a csv file with their number,    #    /||\ |
#              type, and description.                                          #  ./ || \|
#                                                                              #     ||  |
# Dependencies: shiny, tabulapdf, DT, stringr, dplyr, bslib                    #   _/  \_
#                                                                              #             
# Author:      Ryan DePalma                                                    #   
# Version:     1.1 ----- Adding functionality with Microsoft Word docx         #
################################################################################

# ------- Overall Shiny Packages ----------- #
library(shiny)            # Core package for building shiny application
library(DT)               # Renders interactive data tables in UI
library(bslib)            # Provides themes and styling for the app
library(shinycssloaders)  # for loading screen in UI
library(shinyWidgets)     # for range input on UI
library(shinyjs)          # for enabling / disabling features
# ------------ PDF Packages ---------------- #
library(tabulapdf)        # Extracts tabular data from PDF files
library(pdftools)         # for getting information about PDF
# ----------- Word Packages ---------------- #
library(docxtractr)       # for extracting text from Word doc
# ----------- General Packages --------------#
library(stringr)          # Handles string operations and pattern matching
library(dplyr)            # For data manipulation


# Increase file upload limit to 20 MB in case of large PDF
options(shiny.maxRequestSize = 20*1024^2)

# Define the User Interface (UI) #
ui <- fluidPage(
  # Applying bootswatch theme
  theme = bs_theme(bootswatch = "yeti"),
  # Use for disabling in future
  useShinyjs(),
  # styling for Title
  tags$head(
    tags$style(HTML("
      .title-panel {
        background-color: #008cba;
        padding: 20px;
        color: white;
        border-bottom: 2px solid #ccc;
        text-align: center;
      }
    "))
  ),
  
  # styling for Notification
  tags$head(
    tags$style(HTML("
    .shiny-notification {
      position: fixed;
      bottom: 20px !important;
      left: 50% !important;
      transform: translateX(-50%) !important;
      top: auto !important;
      z-index: 9999;
    }
  "))
  ),
  
  # Title Panel
  titlePanel(HTML("<h1 class='title-panel'> TLF EXTRACTOR WIZARD </h1>")),
  # Sidebar for user inputs
  sidebarLayout(
    sidebarPanel(
      div( style = "text-align: left;",
           img(src = "sunLogo1.png", height = "auto", width = "10%")),
      div( style = "text-align: center;", 
           img(src = "wizardhat.png", height = "auto", width = "35%")),
      radioButtons("file_type", "Select File Type:",
                   choices=c("PDF"="pdf","Word"="word"), 
                   inline=TRUE),
      uiOutput("file_input_ui"),   # placeholder for file input
      # panel for inputs for pdf
      conditionalPanel("input.file_type == 'pdf'",
                       numericRangeInput("page_range", "Select Page Range * ",
                                         value = c(1, 10), min = 1, max = 300),
                       HTML("<div style='margin-top: 20px;'></div>"),    # spacing 
                       fluidRow(
                         tags$h5("TLF Number Prefixes"),
                         column(4,
                                textInput("table_prefix", "Table Prefix *", value = "14", width = "100%")
                         ),
                         column(4,
                                textInput("listing_prefix", "Listing Prefix *", value = "16", width = "100%")
                         ),
                         column(4,
                                textInput("figure_prefix", "Figure Prefix *", value = "14", width = "100%")
                         )
                       )
      ),
      HTML("<div style='margin-top: 10px;'></div>"),    # spacing 
      actionButton("extract_btn", "Extract TLF Information", class="btn-info"),
      uiOutput("download_ui"),   # placeholder for download button
      HTML("<div style='margin-top: 15px;'></div>"),    # spacing 
      helpText("Tip: After Clicking 'Extract', please allow some time for processing of PDF and extraction of data.")
    ),
    # Main panel to display data table  
    mainPanel(
      uiOutput("alert_box"),
      withSpinner(
        uiOutput("table_previews"), type = 8),    # preview tables for selection FOR WORD FILES 
      withSpinner(
        DTOutput("tables_output"), type = 8)   # final data table 
    )
  )
)

# Define Server Logic #
server <- function(input, output, session) {
  
  extracted_data <- reactiveVal(NULL)
  tables <- reactiveVal(NULL)
  
  # Dynamically render file input in UI placeholder
  output$file_input_ui <- renderUI({
    if (input$file_type == "pdf") {
      fileInput("pdf_file", "Upload PDF File", accept = ".pdf")
    } else {
      fileInput("word_file", "Upload Word File", accept = ".docx")
    }
  })
  
  # Read Word tables on upload
  observeEvent(input$word_file, {
    req(input$word_file)
    doc <- read_docx(input$word_file$datapath)
    tables(docx_extract_all_tbls(doc, preserve = TRUE))
    # lock the radio button once word file uploaded
    if (!is.null(input$word_file)) {
      disable("file_type")  # Lock the radio buttons
    }
  })
  
  # Render Word table previews
  output$table_previews <- renderUI({
    req(input$file_type == "word", tables())
    lapply(seq_along(tables()), function(i) {
      tagList(
        checkboxInput(paste0("include_table_", i), paste("Include Table", i), value = FALSE),
        tableOutput(paste0("preview_table_", i)),
        tags$hr()
      )
    })
  })
  
  observe({
    req(input$file_type == "word", tables())
    lapply(seq_along(tables()), function(i) {
      output[[paste0("preview_table_", i)]] <- renderTable({
        head(tables()[[i]], 3)
      })
    })
  })
  
  # lock the radio button once pdf file uploaded
  observeEvent(input$pdf_file, { 
    if (!is.null(input$pdf_file)) {
      disable("file_type")
    }
  })
  
  # Main extraction logic
  observeEvent(input$extract_btn, {
    output$alert_box <- renderUI({ NULL })
    
    ### FILE TYPE : PDF .PDF ###
    if (input$file_type == "pdf") {
      # Required files and definitions
      req(input$pdf_file, input$page_range)
      pdf_path <- input$pdf_file$datapath
      page_range <- input$page_range[1]:input$page_range[2]
      
      # Extracting notification
      showNotification("Extracting from PDF...", type = "message")
      
      # Extraction with TABULA PDF's "lattice" method 
      # Lattice is best for well-structured grid tables
      tables_pdf <- extract_tables(pdf_path, pages = page_range, guess = FALSE,
                                   method = "lattice", col_names = FALSE)
      
      # Converting each table row into single line of text for easier parsing
      lines <- unlist(lapply(tables_pdf, function(tbl) {
        apply(tbl, 1, function(row) paste(ifelse(is.na(row), "", row), collapse = " "))
      }))
      
      # Initialize variables
      entries <- list()
      current_entry <- NULL
      current_type <- NULL
      
      # Identify and group lines into Table, Listing, or Figure entries
      for (line in lines) {
        if (str_detect(line, "Table|Figure|Listing")) {
          if (!is.null(current_entry)) {
            entries <- append(entries, list(c(type = current_type, text = current_entry)))
          }
          current_entry <- line
          current_type <- str_extract(line, "Table|Figure|Listing")
        } else if (!is.null(current_entry)) {
          # Stop at "CONFIDENTIAL" (else footer gets added to some entries)
          # & Stop after ".....+"  (else dots in TOC gets added to some entries)
          if (str_detect(line, "CONFIDENTIAL") || str_detect(line, "\\.{3,}\\s*\\d+\\s*$")) {
            current_entry <- paste(current_entry, line)
            entries <- append(entries, list(c(type = current_type, text = current_entry)))
            current_entry <- NULL
            current_type <- NULL
          } else {
            current_entry <- paste(current_entry, line)
          }
        }
      }
      # Ensure last entry is added 
      if (!is.null(current_entry)) {
        entries <- append(entries, list(c(type = current_type, text = current_entry)))
      }
      
      # Extract TLF number and description with Regex
      entries_df <- do.call(rbind, lapply(entries, function(e) {
        type <- e["type"]
        text <- e["text"]
        # Defining Regex based on entry type
        pattern <- switch(type,
                          "Table" = paste0("Table\\s+(", input$table_prefix, "(?:\\.\\d+){1,6})\\s*(.*)"),
                          "Figure" = paste0("Figure\\s+(", input$figure_prefix, "(?:\\.\\d+){1,6})\\s*(.*)"),
                          "Listing" = paste0("Listing\\s+(", input$listing_prefix, "(?:\\.\\d+){1,6})\\s*(.*)")
        )
        
        match <- str_match(text, pattern)
        if (!is.na(match[1])) {
          # Clean up description
          # Note: description needed cleaning because of text formatted TOC
          desc <- str_trim(match[3])
          desc <- str_replace_all(desc, "\\.{3,}\\s*\\d+\\s*", "")
          desc <- str_replace_all(desc, "\\s*\\.{3,}\\s*", "")  # Remove leftover dots
          desc <- str_replace_all(desc, "\\s*\\d+\\s*$", "")    # Remove trailing numbers
          desc <- str_squish(desc)                              # Normalize spacing
          
          data.frame(
            Category = type,
            Number = match[2],
            Description = desc,
            stringsAsFactors = FALSE
          )
        } else {
          NULL
        }
      }))
      # Error Message for PDF or page range missing TLFs
      if (is.null(entries_df) || nrow(entries_df) == 0) {
        output$alert_box <- renderUI({
          tags$div(class = "alert alert-danger", role = "alert",
                   "⚠️ No TLFs found in the selected PDF range. Please check and readjust")
        })
        return(NULL)
      } else {
        output$alert_box <- renderUI({ NULL })  # Clear alert if data is valid
      }
      
      # Data cleaning
      df <- entries_df %>%
        filter(!is.na(Number)) %>%
        distinct() %>%
        select(Number, Category, Description)
      
      # Rename columns for Notion tracker
      colnames(df)[colnames(df) == "Number"] <- "Table Number"
      colnames(df)[colnames(df) == "Description"] <- "Table Title"
      
      # Adding Empty Columns for Notion Tracker 
      df$"SPIL Stat/Programmer Reviewer" <- ""
      df$"SPIL Clinical Reviewer" <- ""
      df$"SPIL Review Comments" <- ""
      df$"'Outside' Stat/Programmer Reviewer" <- ""
      df$"'Outside' Review Date" <- ""
      df$"'Outside' Review Comments" <- ""
      df$"'Outside 2' Comments" <- ""
      
      # Store the result in extracted_data
      extracted_data(df)
      
      ### FILE TYPE : WORD .DOCX ### 
    } else if (input$file_type == "word") {
      
      # Clear table previews in UI
      output$table_previews <- renderUI({ NULL })
      
      # Clear any previous alerts
      output$alert_box <- renderUI({ NULL })
      
      req(input$word_file)
      
      # Get selected tables using isolate to avoid premature reactivity
      included <- isolate(sapply(seq_along(tables()), function(i) input[[paste0("include_table_", i)]] %||% FALSE))
      
      # Check if any tables are selected
      if (!any(included)) {
        output$alert_box <- renderUI({
          tags$div(class = "alert alert-warning", role = "alert",
                   "⚠️ No tables selected from Word document. Refresh site")
        })
        return(NULL)
      }
      # Notice that extraction has started
      showNotification(ui = tags$div(style="font-size:24px; padding:10px;",
                                     "Extraction has started"), type = "message") 
      
      selected_tables <- tables()[which(included)]
      
      # Combine selected tables into one data frame
      df <- do.call(rbind, lapply(selected_tables, function(tbl) tbl))
      
      
      # Error Message for missing TLFs
      if (is.null(df) || nrow(df) == 0) {
        output$alert_box <- renderUI({
          tags$div(class = "alert alert-danger", role = "alert",
                   "⚠️ No data found in selected Word tables.")
        })
        return(NULL)
      } else {
        output$alert_box <- renderUI({ NULL })  # Clear alert if data is valid
      }
      
      colnames(df) <- c("Table Number", "Table Title")
      # Assigning category
      df$Category <- case_when(
        str_detect(tolower(df$`Table Number`), "^f|figure") ~ "Figure",
        str_detect(tolower(df$`Table Number`), "^l|listing") ~ "Listing",
        str_detect(tolower(df$`Table Number`), "^t|table") ~ "Table",
        TRUE ~ "Unknown"
      )
      # Reducing Table number to just its number
      df$`Table Number` <- str_extract(df$`Table Number`, "\\d+(\\.\\d+)*")
      
      # Adding Empty Columns for Notion Tracker 
      df$"SPIL Stat/Programmer Reviewer" <- ""
      df$"SPIL Clinical Reviewer" <- ""
      df$"SPIL Review Comments" <- ""
      df$"'Outside' Stat/Programmer Reviewer" <- ""
      df$"'Outside' Review Date" <- ""
      df$"'Outside' Review Comments" <- ""
      df$"'Outside 2' Comments" <- ""
      
      # Store the result
      extracted_data(df)
    }
  })
  
  # Download button UI
  output$download_ui <- renderUI({
    data <- extracted_data()
    if (!is.null(data) && nrow(data) > 0) {
      downloadButton("downloadData", "Download", class = "btn-primary")
    }
  })
  
  # Configure DataTable display
  output$tables_output <- renderDT({
    req(extracted_data())
    datatable(extracted_data(), options = list(pageLength = 15, autoWidth = TRUE), rownames = FALSE)
  })
  
  # Download handler
  output$downloadData <- downloadHandler(
    filename = function() {
      "exportedTLFinfo.csv"
    },
    content = function(file) {
      write.csv(extracted_data(), file, row.names = FALSE)
    }
  )
}

shinyApp(ui = ui, server = server)#    /\    
#                                  \  /__\
#                                   \ {00}
#                                    \/||\.
#                                     '||  
#                                      ||  
#                                    _/  \_
