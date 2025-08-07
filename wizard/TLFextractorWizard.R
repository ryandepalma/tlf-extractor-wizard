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
# Version:     1.2 --- Adding Excel Functionality .xlsx                        #
################################################################################
# ------- Overall Shiny Packages --------------------------------------------- #
library(shiny)            # Core package for building shiny application        #
library(DT)               # Renders interactive data tables in UI              #
library(bslib)            # Provides themes and styling for the app            #
library(shinycssloaders)  # Shows loading animations in UI                     #
library(shinyWidgets)     # Provides enhanced UI components                    #
library(shinyjs)          # Enables javascript features like disabling inputs  #
# ------------ PDF Packages -------------------------------------------------- #
library(tabulapdf)        # Extracts tabular data from PDF files               #
library(pdftools)         # Reads PDF metadata, text, and page counts          #
# ----------- Word Packages -------------------------------------------------- #
library(docxtractr)       # Extracts tables and text from Word (.docx) files   #
# ------------ Excel Packages ------------------------------------------------ #
library(readxl)           # Extracts tables and text from Excel (.xlsx) files  #
# ----------- General Packages ----------------------------------------------- #
library(stringr)          # Handles string operations and pattern matching     #
library(dplyr)            # For data manipulation                              #
# ---------------------------------------------------------------------------- #
################################################################################

# Global #
# Increase file upload limit to 20 MB in case of large File
options(shiny.maxRequestSize = 20*1024^2)

# Define the User Interface (UI) #
ui <- fluidPage(
  # Applying bootswatch theme
  theme = bs_theme(bootswatch = "yeti"),
  
  # tab title
  tags$head(
    tags$title("TLF Extractor Wizard")
  ),
  
  # styling for Title ; could try switching to bootswatch class
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
  
  # styling for Notifications (location) ; bootswatch classes used in showNotification()
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
  # Use for disabling in future
  useShinyjs(),
  
  # Title Panel
  titlePanel(HTML("<h1 class='title-panel'> TLF EXTRACTOR WIZARD </h1>")),
  
  # Sidebar for user inputs
  sidebarLayout(
    sidebarPanel(
      # sun logo ##########################################################
      div( style = "text-align: left;",                                   # Logos are stored in www folder
           img(src = "sunLogo1.png", height = "auto", width = "10%")),    # Ensure they stay there or remove from UI
      # wizard hat logo ###################################################
      div( style = "text-align: center;", 
           img(src = "wizardhat.png", height = "auto", width = "35%")),
      radioButtons("file_type", "Select File Type:",
                   choices=c("PDF"="pdf","Word"="word","Excel"="excel"), 
                   inline=TRUE), # Radio button for file type 
      helpText("Once file is uploaded, radio button will lock. Refresh site to reset this."),
      uiOutput("file_input_ui"), # Placeholder for file input
      
      # INPUT PANEL FOR PDF FILES
      conditionalPanel("input.file_type == 'pdf' ", 
                       numericRangeInput("page_range", "Select Page Range * ",
                                         value = c(1, 10), min = 1, max = 300), # Page range input for pdfs
                       HTML("<div style='margin-top: 20px;'></div>"), # spacing 
                       fluidRow(
                         tags$h5("TLF Number Prefixes"), # fluid row for prefix inputs
                         column(4,
                                textInput("table_prefix", "Table Prefix *", value = "14", width = "100%")
                         ),
                         column(4,
                                textInput("listing_prefix", "Listing Prefix *", value = "16", width = "100%")
                         ),
                         column(4,
                                textInput("figure_prefix", "Figure Prefix *", value = "14", width = "100%")
                         )
                       ),
                       HTML("<div style='margin-top: 10px;'></div>"), # spacing 
                       actionButton("extract_btn", "🛠Extract TLF Information", class="btn-info")
      ),
      # INPUT PANEL FOR WORD FILES
      conditionalPanel("input.file_type == 'word' ",
                       HTML("<div style='margin-top: 10px;'></div>"), # spacing 
                       actionButton("extract_btn", "🛠Extract TLF Information", class="btn-info"),
      ),
      # INPUT PANEL FOR EXCEL FILES 
      conditionalPanel("input.file_type == 'excel' ", 
                       uiOutput("sheet_selector"), # Placeholder for sheet selector
                       uiOutput("column_selector"), # Placeholder for column selector
                       actionButton("extract_btn", "🛠️Extract TLF Information", class="btn-info"),
      ),
      uiOutput("download_ui"), # Placeholder for download button
      HTML("<div style='margin-top: 15px;'></div>"), # spacing 
      helpText("Tip 1: After Clicking 'Extract', please allow some time for processing of the File and extraction of data.")
    ),
    # Main panel to display data table  
    mainPanel(
      uiOutput("preview_title"),    # placeholder for preview TITLE (for both word and excel)
      uiOutput("preview_instructions"),  # placeholder for instructions for previews (for both word and excel)
      withSpinner(
        uiOutput("table_previews"), type = 8),    # preview tables for selection FOR WORD FILES 
      withSpinner(
        dataTableOutput("sheet_preview"), type = 8),    # preview excel sheets 
      uiOutput("final_title"),   # placeholder for TLF Data Table TITLE
      uiOutput("final_title_instructions"), # placeholder for TLF Data Table instructions
      withSpinner(
        DTOutput("tables_output"), type = 8)   # final data table 
    )
  )
)

# Define Server Logic #
server <- function(input, output, session) {
  
  # Initialize 
  extracted_data <- reactiveVal(NULL)
  tables <- reactiveVal(NULL)
  
  # Dynamically render file input in UI placeholder
  # only allow appropriate file type with accept = 
  output$file_input_ui <- renderUI({
    switch(input$file_type,
           "pdf" = fileInput("pdf_file", "Upload PDF File (.pdf)", accept = ".pdf"),
           "word" = fileInput("word_file", "Upload Word File (.docx)", accept = ".docx"),
           "excel" = fileInput("excel_file", "Upload Excel File (.xlsx)", accept = ".xlsx")
    )
  })
  
  # WORD DEPENDECIES UPON UPLOAD START #
  # Read Word tables on upload
  observeEvent(input$word_file, {
    req(input$word_file)
    doc <- read_docx(input$word_file$datapath)
    tables(docx_extract_all_tbls(doc, preserve = TRUE))
  })
  
  # Render Word table previews in UI mainpanel
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
  # Render Previews Title once word is uploaded
  observe({
    req(input$file_type == "word", input$word_file)
    output$preview_title <- renderUI({
      tags$h3("📊 Preview of Tables in Word Document", class = "bg-info")
    })
    output$preview_instructions <- renderUI({
      tags$p("Please wait for previews of tables to load, as this may take a moment. Once all tables of interest are selected, extract.")
    })
  })
  # Render Table previews in UI mainpanel
  observe({
    req(input$file_type == "word", tables())
    lapply(seq_along(tables()), function(i) {
      output[[paste0("preview_table_", i)]] <- renderTable({
        head(tables()[[i]], 3)
      })
    })
  })     # WORD DEPENDECIES UPON UPLOAD END #
  
  # EXCEL DEPENDECIES UPON UPLOAD START #
  # Reactive storage of of uploaded file path
  excel_path <- reactive({
    req(input$excel_file)
    input$excel_file$datapath
  })
  
  # Reactive Listing in UI sidebar of sheets in excel file
  output$sheet_selector <- renderUI({
    req(excel_path())
    sheets <- readxl::excel_sheets(excel_path())
    selectInput("selected_sheet", "Select Sheet:", choices = sheets)
  })
  # Render Title for Sheet Preview
  observeEvent(input$excel_file, {
    output$preview_title <- renderUI({
      tags$h3(paste0("📊 Sheet Preview of '",input$selected_sheet ,"' "), class = "bg-info")
    })
    output$preview_instructions <- renderUI({
      tags$p("Use this preview to view which columns to select for extraction.")
    })
  })
  # Reactive rendering of selected sheet in UI mainpanel
  sheet_data <- reactive({
    req(excel_path(), input$selected_sheet)
    readxl::read_excel(excel_path(), sheet = input$selected_sheet)
  })
  output$sheet_preview <- renderDataTable({
    req(sheet_data())
    datatable(head(sheet_data(), 10), options = list(scrollX = TRUE))
  })
  
  # Show Column selector in UI sidebar for extraction 
  output$column_selector <- renderUI({
    req(sheet_data())
    cols <- names(sheet_data())  # names of sheets in excel
    tagList(
      selectInput("col1", HTML("Select 'Number' Column: <br> [i.e. has values like '14.1.1']"), choices = cols),
      selectInput("col2", HTML("Select 'Title' Column: <br> [i.e. has values like 'Subject Enrollment by Site – Randomized Subjects']"), choices = cols)
    )
  })    # EXCEL DEPENDECIES UPON UPLOAD END #
  
  # Lock the radio button when file uploaded
  # PDF
  observeEvent(input$pdf_file, { 
    if (!is.null(input$pdf_file)) {
      disable("file_type")
    }
  })
  # WORD
  observeEvent(input$word_file, { 
    if (!is.null(input$word_file)) {
      disable("file_type")
    }
  })
  # EXCEL 
  observeEvent(input$excel_file, { 
    if (!is.null(input$excel_file)) {
      disable("file_type")
    }
  })
  
  # Main extraction logic
  observeEvent(input$extract_btn, {
    
    ### FILE TYPE : PDF .PDF ###
    if (input$file_type == "pdf") {
      # Required files and definitions
      req(input$pdf_file, input$page_range)
      pdf_path <- input$pdf_file$datapath
      page_range <- input$page_range[1]:input$page_range[2]
      
      # Extracting notification
      showNotification("Extracting Tables from PDF file...", type = "message")
      
      # Render Title of Compiled Data Table
      output$final_title <- renderUI({
        tags$h3(paste0("📊 Data Table of Tables, Listings, & Figures"), class = "bg-info")
      })
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
          # Stop at "CONFIDENTIAL" (else footer gets added to some entries)       # This was a problem in a TILD 19-19 TOC 
          # & Stop after ".....+"  (else dots in TOC gets added to some entries)  # This was a problem in all Deurux TOCs
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
          # found this issue with the Deurux Studies TOCs. You can omit this, but otherwise ".... 3" etc was be included in title
          desc <- str_trim(match[3])
          desc <- str_replace_all(desc, "\\.{3,}\\s*\\d+\\s*", "")
          desc <- str_replace_all(desc, "\\s*\\.{3,}\\s*", "")  # Remove leftover dots
          desc <- str_replace_all(desc, "\\s*\\d+\\s*$", "")    # Remove trailing numbers
          desc <- str_squish(desc)                              # Normalize spacing
          
          # Matching columns to data
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
        showNotification("⚠️ No TLFs found in the selected PDF range. Please check and readjust", type = "error")
        return()
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
      output$preview_title <- renderUI ({ NULL })
      output$preview_instructions <- renderUI ({ NULL })
      output$table_previews <- renderUI({ NULL })
      
      req(input$word_file)
      
      # Retrieve the inclusion status (TRUE/FALSE) for each table input, defaulting to FALSE if not set; 
      # isolate to prevent reactive dependencies.
      included <- isolate(sapply(seq_along(tables()), function(i) input[[paste0("include_table_", i)]] %||% FALSE))
      
      # Check if any tables are selected
      if (!any(included)) {
        showNotification("⚠️ No tables selected from Word document. Refresh site", type="error")
        return()
      }
      # Notice that extraction has started
      showNotification(ui = tags$div("Extracting Selected Tables from Word Document..."), type = "message") 
      
      # Render Title of Compiled Data Table
      output$final_title <- renderUI({
        tags$h3(paste0("📊 Data Table of Tables, Listings, & Figures"), class = "bg-info")
      })
      # Grab selected tables from UI
      selected_tables <- tables()[which(included)]
      
      # Combine selected tables into one data frame
      df <- do.call(rbind, lapply(selected_tables, function(tbl) tbl))
      
      # Rename column 1 and 2 appropriately
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
      
      ### FILE TYPE : EXCEL .XLSX ### 
    } else if (input$file_type == "excel") {
      if (is.null(input$excel_file)) {
        showNotification("⚠️ Please upload a Word File before attempting to extract TLF information.", type="error")
        return()
      }
      req(sheet_data(), input$col1, input$col2)
      
      # Notification that extraction has started
      showNotification(paste0("Extracting '",input$selected_sheet ,"' Sheet from Excel file..."), type = "message")
      
      # Extract selected columns and rename them to standard names
      new_data <- sheet_data()[, c(input$col1, input$col2)]
      colnames(new_data) <- c("Number", "Title")
      
      # Clean the 'Number' column to extract numeric patterns
      new_data$Number <- str_extract(new_data$Number, "\\d+(\\.\\d+)*")
      
      # Add a Category column for T, L, or F
      # Category is assigned by which sheet is extracted
      new_data$Category <- input$selected_sheet
      
      # Remove rows with any NA values
      new_data <- na.omit(new_data)
      
      # Remove duplicates within this sheet
      new_data <- distinct(new_data)
      
      # Adding Empty Columns for Notion Tracker 
      new_data$"SPIL Stat/Programmer Reviewer" <- ""
      new_data$"SPIL Clinical Reviewer" <- ""
      new_data$"SPIL Review Comments" <- ""
      new_data$"'Outside' Stat/Programmer Reviewer" <- ""
      new_data$"'Outside' Review Date" <- ""
      new_data$"'Outside' Review Comments" <- ""
      new_data$"'Outside 2' Comments" <- ""
      
      # Combine with existing data
      df <- rbind(extracted_data(), new_data)
      df <- distinct(df)
      
      # Render Title and instructions of Compiled Data Table
      output$final_title <- renderUI({
        tags$h3(paste0("📊 Data Table of Tables, Listings, & Figures"), class = "bg-info")
      })
      output$final_title_instructions <- renderUI({
        tags$p(HTML("Before downloading, ensure you have extracted from each sheet for Tables, Listings, and Figures. <br> 
               Recommendation: Check you have each category by using the search bar to the right and searching its category."))
      })
      
      # Store the result
      extracted_data(df)
    }
  })
  
  # Download button UI
  output$download_ui <- renderUI({
    data <- extracted_data()
    if (!is.null(data) && nrow(data) > 0) {
      downloadButton("downloadData", "Download", class = "btn-primary") # render download button when data is available
    }
  })
  # Configure DataTable display
  output$tables_output <- renderDT({
    req(extracted_data())
    datatable(extracted_data(), options = list(pageLength = 15, autoWidth = TRUE), rownames = FALSE) # render Data Table
  })
  
  # Download handler
  output$downloadData <- downloadHandler(
    filename = function() {
      "notion_TLF_tracker_info.csv" # name of csv to be created
    },
    content = function(file) {
      write.csv(extracted_data(), file, row.names = FALSE)  # writing as csv file
    }
  )
}
# Run App
shinyApp(ui = ui, server = server)#    /\    
#                                  \  /__\
#                                   \ {00}
#                                    \/||\.
#                                     '||  
#                                      ||  
#                                    _/  \_
