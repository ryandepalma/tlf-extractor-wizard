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
# Version:     1.0                                                             #
################################################################################

library(shiny)            # Core package for building shiny application
library(tabulapdf)        # Extracts tabular data from PDF files
library(DT)               # Renders interactive data tables in UI
library(stringr)          # Handles string operations and pattern matching
library(dplyr)            # For data manipulation
library(bslib)            # Provides themes and styling for the app
library(pdftools)         # for getting information about PDF
library(shinycssloaders)  # for loading screen in UI
library(shinyWidgets)     # for range input on UI

# Increase file upload limit to 20 MB in case of large PDF
options(shiny.maxRequestSize = 20*1024^2)

# Define the User Interface (UI) #
ui <- fluidPage(
  # Applying bootswatch theme
  theme = bs_theme(bootswatch = "yeti"),
  
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
  
  # Title Panel
  titlePanel(HTML("<h1 class='title-panel'> TLF EXTRACTOR WIZARD </h1>")),
  # Sidebar for user inputs
  sidebarLayout(
    sidebarPanel(
      div( style = "text-align: left;",
           img(src = "sunLogo1.png", height = "auto", width = "10%")),
      div( style = "text-align: center;", 
           img(src = "wizardhat.png", height = "auto", width = "35%")),
      fileInput("pdf_file", "Upload PDF", accept =".pdf"),                          
      numericRangeInput("page_range", "Select Page Range",
                        value = c(1, 10), min = 1, max = 300),
      HTML("<div style='margin-top: 20px;'></div>"),    # spacing 
      fluidRow(
        tags$h5("TLF Number Prefixes"),
        column(4,
               textInput("table_prefix", "Table Prefix", value = "14", width = "100%")
        ),
        column(4,
               textInput("listing_prefix", "Listing Prefix", value = "16", width = "100%")
        ),
        column(4,
               textInput("figure_prefix", "Figure Prefix", value = "14", width = "100%")
        )
      ),
      HTML("<div style='margin-top: 10px;'></div>"),    # spacing 
      actionButton("extract_btn", "Extract TLF Information", class="btn-info"),
      uiOutput("download_ui"),
      HTML("<div style='margin-top: 15px;'></div>"),    # spacing 
      helpText("Tip: After Clicking 'Extract', please allow some time for processing of PDF and extraction of data.")
    ),
  # Main panel to display data table  
  mainPanel(
      uiOutput("alert_box"),
      withSpinner(
        DTOutput("tables_output"), type = 8)
    )
  )
)

# Define Server Logic #
server <- function(input, output, session) {
  
  # Get number of pages in uploaded PDF
  num_pages <- reactive({
    req(input$pdf_file)
    info <- pdf_info(input$pdf_file$datapath)
    info$pages
  })
  
  # Store extracted data for use in other observers
  extracted_data <- reactiveVal(NULL)
  
  # Extracting Data once button selected 
  observeEvent(input$extract_btn, {
    
    # Check if PDF is uploaded
    if (is.null(input$pdf_file)) {
      output$alert_box <- renderUI({
        tags$div(class = "alert alert-danger", role = "alert",
                 "⚠️ Please upload a PDF before attempting to extract TLF information.")
      })
      return()
    }
    # Clear any previous alerts if file is present
    output$alert_box <- renderUI({ NULL })
    
    # Ensure required inputs are available before proceeding 
    req(input$pdf_file)       
    req(input$page_range)
    
    # Define file path and generate sequence of page range
    pdf_path <- input$pdf_file$datapath     
    page_range <- input$page_range[1]:input$page_range[2]  
    
    # Extraction with TABULA PDF's "lattice" method 
    # Lattice is best for well-structured grid tables
    tables <- extract_tables(pdf_path, pages = page_range, guess = FALSE,
                              method = "lattice", col_names = FALSE)

    # Converting each table row into single line of text for easier parsing
    lines <- unlist(lapply(tables, function(tbl) {
      apply(tbl, 1, function(row) paste(ifelse(is.na(row), "", row), collapse = " "))
    }))
    
    # Initialize variables
    entries <- list()
    current_entry <- NULL 
    current_type <- NULL
    
    # Identify and group lines into Table, Listing, or Figure entries
    for (line in lines) {
      if (str_detect(line, "Table")) {
        if (!is.null(current_entry)) {
          entries <- append(entries, list(c(type = current_type, text = current_entry)))
        }
        current_entry <- line
        current_type <- "Table"
      } else if (str_detect(line, "Figure")) {
        if (!is.null(current_entry)) {
          entries <- append(entries, list(c(type = current_type, text = current_entry)))
        }
        current_entry <- line
        current_type <- "Figure"
      } else if (str_detect(line, "Listing")) {
        if (!is.null(current_entry)) {
          entries <- append(entries, list(c(type = current_type, text = current_entry)))
        }
        current_entry <- line
        current_type <- "Listing"
      } else if (!is.null(current_entry)) {
        # Stop at "CONFIDENTIAL" (else footer gets added to some entries)
        if (str_detect(line, "CONFIDENTIAL")) {
          entries <- append(entries, list(c(type = current_type, text = current_entry)))
          current_entry <- NULL
          current_type <- NULL
        } 
        # Stop after ".....+"
        else if (str_detect(line, "\\.{3,}\\s*\\d+\\s*$")) {
          current_entry <- paste(current_entry, line)  # Include the final line
          entries <- append(entries, list(c(type = current_type, text = current_entry)))
          current_entry <- NULL
          current_type <- NULL
        }
        else {
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
                 "⚠️ No Tables, Listings, or Figures were detected in the selected page range. Please try a different page range or PDF.")
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
    return(df)
  })
  
  # Download button to appear when conditions are met for export
  output$download_ui <- renderUI({
    data <- extracted_data()
    if (!is.null(data) && nrow(data) > 0) {
      downloadButton("downloadData", "Download", class = "btn-primary")
    } else {
      NULL
    }
  })
  
  # Configure DataTable display
  output$tables_output <- renderDT({
    req(extracted_data())
    datatable(extracted_data(), options = list(pageLength = 15, autoWidth = TRUE), rownames = FALSE)
  })
  
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