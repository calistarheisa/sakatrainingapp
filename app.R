library(shiny)
library(shinydashboard)
library(shinythemes)
library(DT)
library(dplyr)
library(lubridate)
library(readxl)
library(stringr)
library(ggplot2)
library(treemapify)
library(tidyr)
library(writexl)
library(RColorBrewer)
library(zoo)
library(shinyjs)
library(coin)
library(stats)
library(tidytext)
library(gt)
library(officer)
library(flextable)
library(rvg)
library(plotly)
library(viridis)
library(packrat)
library(rsconnect)
library(openxlsx)
library(renv)

# Dataset
# getwd()
# setwd("C:/Users/Calista/OneDrive/Documents/KP PGN Saka/app")

raw_data <- read_excel("raw_data_dummy.xlsx")
raw_data;
as.character(raw_data$exp_date)
as.character(raw_data$start_date)
as.character(raw_data$end_date)

# Header
header <- dashboardHeader(
  title = "Training Dashboard",
  titleWidth = 300,
  dropdownMenu(
    headerText = "Meet the Creators!",
    type = "messages",
    icon = icon("user"),
    messageItem(
      from = "Rosita Adelia Puspitasari",
      message = "Click here!",
      icon = icon("linkedin"),
      href = "https://www.linkedin.com/in/rositaadeliapuspitasari/"
    ),
    messageItem(
      from = "Calista Fara Rheisa",
      message = "Click Here!",
      icon = icon("linkedin"),
      href = "https://www.linkedin.com/in/calistarheisa/"
    )
  )
)

# Sidebar
sidebar <- dashboardSidebar(
  width = 200,
  sidebarMenu(
    menuItem("Home", tabName = "Home", icon = icon("home")),
    menuItem("Training Record", tabName = 'train_record', icon = icon("book"),
             menuSubItem("Dashboard", tabName = "Dashboard", icon = icon("dashboard")),
             menuSubItem("Record", tabName = "Record", icon = icon("user")),
             menuSubItem("Training Details", tabName = "Course", icon = icon("book")),
             menuSubItem("Input Data", tabName = "Input",icon = icon("edit")),
             menuSubItem("View and Edit Data", tabName = "Data",icon = icon("table")),
             menuSubItem("Export Data", tabName = "Export",icon = icon("database"))),
    menuItem("Training Evaluation", tabName = "evaluation", icon = icon("table"),
             menuSubItem("Data", tabName = "data", icon = icon("upload")),
             menuSubItem("Report", tabName = "report", icon = icon("clipboard"))),
    menuItem("Creators", tabName = "creators", icon = icon("user"))
    ))

# Body
body <- dashboardBody(
  tags$head(
    tags$style(HTML("
    body {
      background-color: #fefefe;
    }
    .main-header {background-color: #134a6e !important;
        }

        /* Ubah warna background logo/title */
    .main-header .logo {
          background-color: #134a6e !important; 
          color: white !important;
          font-size: 20px !important;
          font-family: 'Trebuchet MS' !important;
          font-weight: bold;
        }

        /* Ubah warna background navbar (kanan header) */
    .main-header .navbar {
          background-color: #134a6e !important
        }

        /* Ubah warna tulisan di navbar */
    .main-header .navbar .nav>li>a {
          color: white !important;
          font-size: 16px !important;
        }
        
    .skin-blue .main-header { background-color: #134a6e; }
    .skin-blue .main-sidebar { background-color: #134a6e; }
    .skin-blue .sidebar-menu > li > a { color: #fff; }
    .skin-blue .sidebar-menu > li > a:hover { background-color: #70bacf; }
    .skin-blue .sidebar-menu > li.active > a { background-color: #70bacf; color: #fff; }
    .skin-blue .sidebar-header { background-color: #134a6e; }
    .skin-blue .sidebar-toggle { background-color: #134a6e; }
    /* Hover effect */
    .sidebar-menu > li > a:hover {
      color: #FFFFFF; /* Teks putih */
      background-color: #70bacf; /* Biru sedang */
    }
    /* Active menu item */
    .sidebar-menu > li.active > a {
      color: #FFFFFF; /* Teks putih */
      background-color: #70bacf; /* Biru terang */
      font-weight: bold;
    }
    
    /* Submenu styles */
    .sidebar-menu .treeview-menu {
      background-color: #6F8FC1; /* Warna latar belakang submenu */
    }
    
    .sidebar-menu .treeview-menu > li > a {
      color: #D8EFFF; /* Warna teks submenu */
      font-size: 13px;
    }
    
    .sidebar-menu .treeview-menu > li > a:hover {
      color: #FFFFFF;
      background-color: #0093ad;
      border-radius: 8px;
    }
    
    /* Notification badge */
    .notification-badge {
      background-color: #0093ad; /* Biru terang */
      color: white;
      font-size: 12px;
      border-radius: 50%;
      padding: 4px 8px;
      text-align: center;
      margin-left: 5px;
    }
    .box-custom {
      background-color: #ffffff;
      border-radius: 8px;
      padding: 20px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.1);
      width: 100%;
      max-width: 1200px;
    }
    .action-btn {
      font-size: 12px; 
      padding: 5px 10px; 
      height: 30px; 
      margin-top: 10px; 
      background-color: #6F8FC1; !important;
      color: white; 
      border: none; 
      border-radius: 5px !important;
    }
    
    #shiny-tab-report {
    background-color: #f0f0f0 !important;
    min-height: calc(100vh - 80px);
    padding: 30px;
    }
    
    .box .box-title {
      text-align: center;
      font-weight: bold;
    }
    .table-custom th, .table-custom td {
      padding: 12px 15px;
      text-align: center;
    }
    .table-custom {
      border-collapse: collapse;
      width: 100%;
      margin: 0 auto
    }
    .table-custom tr:nth-child(even) {
      background-color: #f2f2f2;
    }
    .container {
      display: flex;
      justify-content: center;  /* Memusatkan konten secara horizontal */
      align-items: center;  /* Memusatkan konten secara vertikal */
      height: 100vh;
    }
    /* Responsiveness */
    @media (max-width: 768px) {
      .employee-stat {
        flex-direction: column;
      }
      .action-btn {
        width: 100%;
        padding: 12px;
      }
    }
  .custom-box {
    background-color: #ffffff !important; /* Warna latar putih */
    color: #1F2855 !important; /* Warna teks */
    border-radius: 10px; /* Membulatkan sudut */
    text-align: center; /* Meratakan semua konten di tengah */
    display: flex;
    flex-direction: column;
    justify-content: center;
    height: 100%; /* Pastikan konten mengikuti tinggi box */
    line-height:1.1;
  }
  .custom-pie plotOutput {
    margin: 0px !important;  # Pastikan margin antar chart tidak ada
    padding: 0px !important; # Hilangkan padding bawaan
  }
  .custom-box > .box-header {
    background-color: #134a6e !important;
    color: #ffffff !important;
    border-top-left-radius: 10px; 
    border-top-right-radius: 10px; 
    text-align: center;
    font-size : 18.5px;
    justify-content: center;
    text-align: center;
  }
  .custom-text {
    font-family : 'Trebuchet MS';
    font-size: 16px; /* Ukuran default untuk teks */
    line-height: 1.1;
    font-weight: bold;
    color: #333333;
  }
  .custom-text .highlight {
    font-size: 33px; /* Ukuran lebih besar untuk nama */
    font-weight: bold; /* Menebalkan nama */
    display: block; /* Pastikan setiap elemen pada baris baru */
    line-height:1.1;
  }
  .box.box-warning {
    background-color: #ffffff !important; /* Biru */
    border-color: #ffffff !important; /* Biru lebih gelap untuk border */
    color: white !important; /* Warna teks */
  }
  .box.box-warning > .box-header {
    background-color: #134a6e !important; /* Warna biru untuk header */
    color: white !important; /* Warna teks header */
    border-top-left-radius: 10px; /* Membulatkan sudut atas kiri */
    border-top-right-radius: 10px; /* Membulatkan sudut atas kanan */
  }
  .fa-trophy {
    font-size: 15px;
    color: gold;
  }
  ")),
    tags$link(rel = "stylesheet", type = "text/css", href = "custom.css")
  ),
  tabItems(
    tabItem(
      tabName = "Home",
      
      div(
        style = "align-items: right;",
        imageOutput("pgn1"),
        style = "width: 150px; height: 50px;"
      ),
      
      div(
        style = "text-align: center; margin-bottom: 20px;",
        h2("Hello, User!", style = "font-weight: bold; color: #134a6e; font-size: 60px"),
        h4("Welcome to PGN Saka Training Dashboard", style = "color: #555; font-size: 35px")
      ),
      
      div(
        style = "display: flex; justify-content: center; gap: 20px;",
        
        # Card 1
        div(
          class = "card",
          style = "width: 25%; padding: 20px; border: 1px solid #dee2e6; border-radius: 10px; background-color: #ffffff; text-align: center;",
          h5("Expired This Month", style = "font-weight: bold; margin-bottom: 10px; font-size: 25px;"),
          h3(style = "font-size: 90px; font-weight: bold; margin: 0; color: #0093ad;", textOutput("expired_this_month")),
          actionButton(
            "btn_expired_this_month", 
            "Click for Detail", 
            class = "action-btn",
            style = "font-size: 12px; padding: 5px 10px; height: 30px; margin-top: 10px; background-color: #0093ad; color: white; border: none; border-radius: 5px;"
          )
        ),
        
        # Card 2
        div(
          class = "card",
          style = "width: 25%; padding: 20px; border: 1px solid #dee2e6; border-radius: 10px; background-color: #ffffff; text-align: center;",
          h5("Expired Next 3 Months", style = "font-weight: bold; margin-bottom: 10px; font-size: 25px;"),
          h3(style = "font-size: 90px; font-weight: bold; margin: 0; color: #006b89;", textOutput("expired_3_months")),
          actionButton(
            "btn_expired_3_months", 
            "Click for Detail", 
            class = "action-btn",
            style = "font-size: 12px; padding: 5px 10px; height: 30px; margin-top: 10px; background-color: #006b89; color: white; border: none; border-radius: 5px;"
          )
        ),
        
        # Card 3
        div(
          class = "card",
          style = "width: 25%; padding: 20px; border: 1px solid #dee2e6; border-radius: 10px; background-color: #ffffff; text-align: center;",
          h5("Expired This Year", style = "font-weight: bold; margin-bottom: 10px; font-size: 25px"),
          h3(style = "font-size: 90px; font-weight: bold; margin: 0; color: #134a6e;", textOutput("expired_this_year")),
          actionButton(
            "btn_expired_this_year", 
            "Click for Detail", 
            class = "action-btn",
            style = "font-size: 12px; padding: 5px 10px; height: 30px; margin-top: 10px; background-color: #134a6e; color: white; border: none; border-radius: 5px;"
          )
        )
      ),
      
      # Certification Detail Section
      div(
        style = "display: flex; align-items: left;",
        h3("Certification Detail:", style = "margin-right: 5px; margin-top: 30px;"),
        h3(textOutput("expired_title"), style = "margin-top: 30px; color: #134a6e; font-weight: bold;")
      ),
      div(
        class = "box-custom",
        style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
        tableOutput("cert_details"),
        style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
        #style = "padding: 20px; border: 1px solid #dee2e6; border-radius: 10px; background-color: #f8f9fa; margin-top: 10px;"
      )
    ),
    tabItem(
      tabName = "Dashboard",
      tags$div(
        style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 15px;",
        imageOutput("pgn2", inline = TRUE),
        tags$div(
          style = "text-align: right",
          h2("Training Record Dashboard", style = "font-weight: bold; font-size: 50px; margin-bottom: 2px; color: #134a6e;"),
          p("Select a year to view data and insights."),
        )
      ),
      div(
        style = "display: flex; justify-content: flex-end; width: 100%;",  # Flex untuk posisi kanan
        div(
          selectInput(
            inputId = "year_choice", 
            label = NULL, 
            choices = NULL
          )
        )
      ),
      fluidRow(
        box(
          width = 3,
          title = tagList(icon('trophy'),"Most Man Hours"), status = "warning", solidHeader = TRUE,
          height='157px',class='custom-box',
          htmlOutput("best_employee_hours",class='custom-text')
        ),
        box(
          width = 3,
          title = tagList(icon('trophy'), "Most Attended Training"), status = "warning", solidHeader = TRUE,
          height='157px',class='custom-box',
          htmlOutput("best_employee_titles",class = "custom-text")
        ),
        box(
          title = "Summary", width = 6, status = "warning", solidHeader = TRUE,
          class = 'custom-box',
          div(
            style = "display: flex; justify-content: space-between; align-items: center; gap: 10px;",  # Flexbox untuk distribusi elemen
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;",  # Membagi ruang secara proporsional
              uiOutput("total_hours")
            ),
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;;",  # Membagi ruang secara proporsional
              uiOutput("total_days")
            ),
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;",  # Membagi ruang secara proporsional
              uiOutput("total_learn_hours") 
            )
          )
        )
      ),
      fluidRow(
        box(
          width = 12,
          title = "Department Treemap", status = "warning", solidHeader = TRUE,
          class = 'custom-box',
          plotOutput("treemap")
        )
      ),
      fluidRow(
        box(
          width = 8,
          title = "Event Categories", status = "warning", solidHeader = TRUE,
          class = 'custom-pie',
          height = "300px",  # Mengurangi ketinggian box
          div(
            style = "display: flex; justify-content: space-between; align-items: center; height: 100%;",  # Membagi ruang secara merata
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;",  # Membagi width secara proporsional
              plotlyOutput("event_cat_pie", height = "214px", width = "214px")
            ),
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;",  # Membagi width secara proporsional
              plotlyOutput("type1_pie", height = "214px", width = "214px")  # Tinggi dan lebar seragam
            ),
            tags$div(
              style = "flex: 1; display: flex; justify-content: center;",  # Membagi width secara proporsional
              plotlyOutput("type2_pie", height = "214px", width = "214px")  # Tinggi dan lebar seragam
            )
          )
        ),
        box(
          width = 4,
          title = "Top 5 Venues", status = "warning", solidHeader = TRUE,
          class='custom-box',
          plotOutput("top_venues_bar", height = "234px", width = "100%")
        )
      ),
      fluidRow(
        box(
          width = 12,
          title = textOutput('chart_title'), status = "warning", solidHeader = TRUE,
          class = 'custom-box',
          plotOutput("man_hours_line_chart", width="100%", height='400px')
        )
      )
    ),
    tabItem(tabName = "Record",
            tags$div(
              style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 0px;",
              imageOutput("pgn3", inline = TRUE),
              tags$div(
                style = "text-align: right",
                h2("Employee Training Record", style = "font-weight: bold; font-size: 50px; margin-bottom: 2px; color: #134a6e;"),
                p("Select a name and year to view the employee training history.")
              )
            ),
            div(
              style = "display: flex; gap: 10px; justify-content: flex-end; align-items: center;",
              selectizeInput(
                inputId = "search_name",
                label = "Enter Employee Name:",
                choices = NULL,    # Akan diperbarui di server
                options = list(
                  placeholder = "Type a name...",
                  maxOptions = 10   # Maksimum opsi yang ditampilkan
                ),
                width = '70%'
              ),
              selectInput(
                inputId = "filter_year",
                label = "Select Year:",
                choices = NULL,   # Akan diperbarui di server
                selected = "All",
                width = '15%'
              ),
              downloadButton("export_button", "Download Training Report", style = 'height: 35px; margin-top: 10px;')
            ),
            fluidRow(
              column(
                width = 4,
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #134a6e; margin-bottom: 20px;",
                  h4("Profile", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px; color: #ffffff;"),
                  uiOutput("employee_details")                  
                )
              ),
              column(
                width = 8,
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #ffffff; margin-bottom: 20px;",
                  h4("Training Statistics", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px;"),
                  uiOutput("training_stats")
                ),
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #ffffff;",
                  h4("Training History", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px;"),
                  uiOutput("training_history")               
                )
              )
            )),
    tabItem(tabName = "Course",
            tags$div(
              style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 15px;",
              imageOutput("pgn4", inline = TRUE),
              tags$div(
                style = "text-align: left",
                h2("Training Details", style = "font-weight: bold; font-size: 50px; margin-bottom: 5px; color: #134a6e;"),
                "Select the training title and year to view the training details."
              )
            ),
            div(
              style = "display: flex; gap: 10px; justify-content: flex-end; align-items: center;",
              selectizeInput(
                inputId = "search_title",
                label = "Enter Training Title:",
                choices = NULL,
                options = list(
                  placeholder = "Type a course title...",
                  maxOptions = 10   # Maksimum opsi yang ditampilkan
                ),
                width = '70%'
              ),
              selectInput(
                inputId = "title_year",
                label = "Select Year:",
                choices = NULL,   # Akan diperbarui di server
                width = '15%'
              ),
              downloadButton("export_participants", "Download Training Details", style = 'height: 35px; margin-top: 10px;')
            ),
            fluidRow(
              column(
                width = 12,
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #134a6e; margin-bottom: 20px;",
                  h4("Course Details", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px; color: #ffffff;"),
                  uiOutput("course_details")                  
                )
              )
            ),
            fluidRow(
              column(
                width = 4,
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #ffffff; margin-bottom: 20px;",
                  h4("Course Statistics", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px;"),
                  uiOutput("course_stats")
                )
              ),
              column(
                width = 8,
                div(
                  class = "card",
                  style = "width: 100%; padding: 20px; border: 1px solid #dee2e6; border-radius: 20px; background-color: #ffffff;",
                  h4("Course Participants", style = "font-weight: bold; font-size: 25px; margin-bottom: 20px;"),
                  tableOutput("course_participants")               
                )
              )
            )
    ),
    tabItem(tabName = "Input",
            tags$div(
              style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 15px;",
              imageOutput("pgn5", inline = TRUE),
              tags$div(
                style = "text-align: right",
                h2("Input Data", style = "font-weight: bold; font-size: 50px; margin-bottom: 2px; color: #134a6e;"),
                p("This page allows you to input new data.")
              )
            ),
            div(
              style="margin-bottom: 0px;",
              h4("File Upload", style = "font-weight: bold; font-size: 20px; margin-bottom: 2px;"),
              fileInput("file_upload","Upload Excel File", accept = c(".xlsx"), width = "100%")
            ),
            div(
              h4("Manual Upload", style = "font-weight: bold; font-size: 20px; margin-bottom: 2px; margin-top: 2px;"),
              textInput("name", "Name:", "",, width = "100%"),
              selectInput(
                "depart", "Department:",
                choices = c(
                  'Corporate Finance', 'SCM', 'Human Resources', 'IT & Data Management', 'Exploration',
                  'Drilling & Well Intervention', 'HSSE', 'Production Muriah', 'Production', 
                  'Facility Engineering', 'Commercial', 'Audit & Ethic Compliance', 'Subsurface',
                  'Project Delivery', 'Operational Finance', 'Legal', 'New Venture',
                  'Planning Strategic Management', 'NOA Management', 'Tax', 'Risk Management',
                  'Stakeholder Relations', 'Corporate Accounting', 'BOD Office', 'Production Pangkah', 
                  'BOC', 'HR & GA', 'Legal & Counsel', 'Finance & Business Support', 'Operations'
                ),
                width = "100%"
              ),
              selectInput(
                'emp_status', 'Employee Status:',
                choices = c('Permanent', 'Contractor', 'Direct Contract', 'Services', 'Magenta Internship'),
                width = "100%"
              ),
              selectInput(
                'work_loc', 'Work Location:',
                choices = c('Jakarta', 'Muriah', 'Gresik', 'Gresik Offshore', 'Gresik Onshore'),
                width = "100%"
              ),
              textInput('cost', 'Cost Center for Training Fee:', "", width = "100%"),
              textInput('title', 'Course Title:', "", width = "100%"),
              dateInput("start_date", "Start Date Training:", startview="year", value = "2000-01-01", width = "100%"),
              dateInput("end_date", "End Date Training:", startview="year", value = "2000-01-01", width = "100%"),
              numericInput("m_days", "Man Days:", value = 0, min = 0, width = "100%"),
              numericInput("m_hours", "Man Hours:", value = 0, min = 0, width = "100%"),
              numericInput("learn_hours", "Learning Hour (LH):", value = 0, min = 0, width = "100%"),
              textInput('venue', 'Venue:', '', width = "100%"),
              textInput('vendor', 'Vendor/Trainer:', '', width = "100%"),
              selectInput('type1', 'Type 1:', choices = c('Certification', 'Technical', 'Non-Technical'), width = "100%"),
              selectInput('event_cat', 'Event Category:', choices = c('Domestic', 'Inhouse', 'Overseas'), width = "100%"),
              selectInput('type2', 'Type 2:', choices = c('Non-HSE', 'HSE'), width = "100%"),
              dateInput('exp_date', 'Expired Date:', value = NULL, width = "100%"),
              actionButton("add_data", "Add Data", class = "btn-primary"),
              br(), br(),
              DTOutput("preview_new_data")
            )
    ),
    
    # Edit Data Tab
    tabItem(tabName = "Data",
            tags$div(
              style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 15px;",
              imageOutput("pgn6", inline = TRUE),
              tags$div(
                style = "text-align: right",
                h2("View and Edit Data", style = "font-weight: bold; font-size: 50px; margin-bottom: 2px; color: #134a6e;"),
                p("This page diplays data and allows you to edit existing data.")
              )
            ),
            useShinyjs(),
            DTOutput("edit_table"),
            actionButton("save_changes", "Save Changes", class = "btn-success"),
            br(), br(),
            verbatimTextOutput("edit_table_cell_edit") 
    ),
    
    # Data View Tab
    tabItem(tabName = "Export",
            tags$div(
              style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 15px;",
              imageOutput("pgn7", inline = TRUE),
              tags$div(
                style = "text-align: right",
                h2("Export Data", style = "font-weight: bold; font-size: 50px; margin-bottom: 2px; color: #134a6e;"),
                p("This page allows you to view and export data in Excel or CSV format.")
              )
            ),
            DTOutput("data_table")
    ),
    
    ####EVALUATION
    
    #Data Upload Tab Evaluation
    tabItem(
      tabName = "data",
      tags$div(
        style = "display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 10px;",
        imageOutput("pgn8", inline = TRUE),
        tags$div(
          style = "text-align: right",
          h4("Training Evaluation Automation Tool", style = "font-weight: bold; font-size: 50px; color: #134a6e;")
        )
      ),
      fluidRow(
        # Kolom pertama untuk informasi dan upload
        column( width = 9,
                div(HTML(
                  "<strong>Training Evaluation Automation Tool</strong> merupakan alat untuk melakukan otomatisasi pelaporan hasil evaluasi pelatihan. Data evaluasi hasil pelatihan yang akan diunggah harus memenuhi beberapa persyaratan sebagai berikut:<br>1. Terdapat 3 data yang perlu diunggah: Data Feedback, Data Pre-Test, dan Data Post-Test.<br>2. Setiap data harus terdiri atas kolom-kolom berikut:<br>• <strong>Data Feedback</strong>: Timestamp, Judul Pelatihan, Tempat Pelaksanaan, Nama Peserta, Departemen, Kejelasan Tujuan, Ketercapaian Tujuan, Kesesuaian Materi,Penerapan Materi, Tambahan Pengetahuan, Tingkat Kesulitan, Kemudahan Pemahaman, Fasilitator Menguasai, Kompetensi Fasilitator, Keaktifan Fasilitator, Penilaian Fasilitator, Total Penilaian, Point Utama Materi, Kritik dan Saran<br>• <strong>Data Pre-Test</strong>: Nama Peserta, Departemen, Nilai Pre-Test<br>• <strong>Data Post-Test</strong>: Nama Peserta, Departemen, Nilai Post-Test<br><em>Note: Format penamaan setiap kolom tidak harus seperti apa yang telah disebutkan.</em><br>3. Penulisan nama karyawan pada ketiga data diusahakan memiliki penulisan yang sama.<br>4. Setiap data yang diunggah harus dalam format <strong>.xlsx</strong>."
                ),
                style = 'font-size: 14px; line-height: 2; white-space: pre-wrap; width: 100%; color: #333; border: 1px solid #ddd; border-radius: 8px; padding: 15px; background-color: #F9F9F9; margin-bottom: 20px; text-align: justify;'
                ),
                # Upload files section
                fluidRow(column(
                  width = 12,
                  h4("Upload Files", style = "font-weight: bold; color: #0056A1; margin-top:0px")
                ),
                column(
                  width = 3,
                  fileInput("feedback_file", "Upload Feedback File", accept = c(".xlsx"), 
                            buttonLabel = "Browse", placeholder = "No file selected")
                ),
                column(
                  width = 3,
                  fileInput("pretest_file", "Upload Pre-Test File", accept = c(".xlsx"), 
                            buttonLabel = "Browse", placeholder = "No file selected")
                ),
                column(
                  width = 3,
                  fileInput("posttest_file", "Upload Post-Test File", accept = c(".xlsx"), 
                            buttonLabel = "Browse", placeholder = "No file selected")
                )
                ),
                # Upload status section
                fluidRow(column(
                  width = 12,
                  uiOutput("upload_status")  # Display message with dynamic color
                ))),
        column(
          width = 3,
          div(
            imageOutput("saka"),
            style = 'width: 100%; height: 100%; border: 1px solid #ddd; border-radius: 8px; padding: 10px; background-color: #F9F9F9;display: flex; align-items: stretch;'
          )))),
    
    # Report Tab
    tabItem(
      tabName = "report",
      fluidRow(
        style = "background-color: #f0f0f0; display: flex; flex-direction: column; align-items: center; text-align: center; padding: 20px;",
        column(
          width = 12,
          div(
            class = "box-custom",
            style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
            fluidRow(
              column(
                width = 2,
                div(
                  imageOutput("pgn"),
                  style = "width: 200px; height: 100px; margin: 10px auto;"
                )
              ),
              column(
                width = 10,
                h3(
                  style = "font-size: 30px; color: #333; margin-bottom: 0px;", 
                  "Training Evaluation"
                ),
                htmlOutput("training_title", style = "margin-bottom: 0px;"),  
                htmlOutput("training_date", style = "margin-bottom: 0px;"),   
                htmlOutput("training_location", style = "margin-bottom: 5px;")
              )
            )
          ),
          h4("Level 1: Training Feedback Visualization", 
             style = "font-weight: bold; color: #333; margin: 20px; font-size:40px; text-align:center;"),
          div(
            class = "box-custom",
            style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
            h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Training Score Bar Chart"),
            plotOutput("training_fig", height = "400px")
          ),
          div(
            class = "box-custom",
            style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
            h3(style = "font-size: 30px; color: #000; font-weight: bold; margin-top: 0px; text-align: left;", "Fasilitator Score Bar Chart"),
            plotOutput("fasil_fig", height = "400px")
          ),
          fluidRow(
            column(
              width = 6,
              div(
                class = "box-custom",
                style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
                h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Difficulty Level Pie Chart"),
                plotOutput("pie_chart", height = "300px")
              )
            ),
            column(
              width = 6,
              div(
                class = "box-custom",
                style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
                h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Facilitator Assessment Pie Chart"),
                plotOutput("score_pie_chart", height = "300px")
              )
            )
          ),
          div(
            class = "box-custom",
            style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
            h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Total Training Score Bar Chart"),
            plotOutput("total_score_fig", height = "400px")
          ),
          fluidRow(
            column(
              width = 6,
              div(
                class = "box-custom",  
                style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
                h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Point Utama Materi"),
                uiOutput("material_summary")
              )
            ),
            column(
              width = 6,
              div(
                class = "box-custom",
                style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
                h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; text-align: left;", "Kritik dan Saran"),
                uiOutput("suggestion_summary")
              )
            )
          ),
          hr(),
          h4("Level 2: Pre-Test and Post-Test Analysis", 
             style = "font-weight: bold; color: #333; margin-top: 20px; font-size:40px; text-align:center;"),
          uiOutput("table_styles"),
          fluidRow(
            column(
              width = 12,
              tableOutput("pre_post_table"),
              style = "margin: 20px; text-align: center; width: 100%;"
            )
          ),
          h2("Treshold Point for Post-Test = 70", 
             style = "margin-bottom: 20px; margin-top: 0px; font-size:15px; text-align:left;"),
          h4("Hasil Analisis", 
             style = "font-weight: bold; color: #333; margin: 20px; font-size:30px; text-align:center;"),
          fluidRow(
            column(
              width = 6,
              div(
                class = "box-custom",
                style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
                h3(style = "font-size: 30px; color: #333; font-weight: bold; margin-top: 0px; margin-bottom: 5px;", "Box Plot Pre-Test vs Post-Test"),
                plotOutput("boxplot", height = "400px")
              )
            ),
            column(
              width = 6,
              div(
                uiOutput("stat_result", style = "margin: 20px;"),
                uiOutput("corr_result", style = "margin: 20px;")
              )
            )
          ),
          downloadButton("download_button", "Download Training Evaluation Report"),
        )
      )
    ),
    tabItem(tabName = "creators",
            div(
              h4("PENYUSUN",style = "font-weight: bold; font-size: 40px; color: #134a6e;"), 
              style = "text-align: center; margin-bottom: 30px;"
            ),
            
            div(
              style = "display: flex; justify-content: center; align-items: center; gap: 20px;",
              div(
                class = "box-custom",
                style = "width: 33rem; height: 44rem;",  # Set width untuk card
                div(
                  imageOutput("calista"), 
                  style = "text-align: center; margin-bottom: -75px;"
                ),
                div(strong("Calista Fara Rheisa"), style = "text-align: center; font-size: 20px; margin-top: 15px;"),
                div("Mahasiswa Kerja Praktik", style = "text-align: center; font-size: 18px;")
              ),
              div(
                class = "box-custom",
                style = "width: 33rem; height: 44rem;",  # Set width untuk card
                div(
                  imageOutput("rosita"), 
                  style = "text-align: center; margin-bottom: -75px;"
                ),
                div(strong("Rosita Adelia Puspitasari"), style = "text-align: center; font-size: 20px; margin-top: 15px;"),  
                div("Mahasiswa Kerja Praktik", style = "text-align: center; font-size: 18px;")
              )
            ),
            div(
              class = "card",
              style = "text-align: center; padding: 20px; margin-top: 20px;",
              div(strong("Program Studi Sarjana Statistika"),style="text-align: center; font-size: 20px;"),
              div(strong("Departemen Statistika"),style="text-align: center; font-size: 20px;"),
              div(strong("Fakultas Sains dan Analitika Data"),style="text-align: center; font-size: 20px"),
              div(strong("Institut Teknologi Sepuluh Nopember Surabaya"),style="text-align: center; font-size: 20px"),
              div(strong("2025"),style="text-align: center; font-size: 20px")
            )
    )
  )
)

ui <- dashboardPage(header = header, body = body, sidebar = sidebar)

# Server Logic
server <- function(input, output, session) {
  
  # Inisialisasi
  rv <- reactiveValues(
    data = read_excel("raw_data_dummy.xlsx")
  )
  
  # Membuat raw_data menjadi referensi dari rv$data
  raw_data <- reactive({
    rv$data  # raw_data selalu sama dengan rv$data yang terbaru
  })
  
  # Reactive untuk mengonversi kolom exp_date menjadi Date di dalam server
  processed_data <- reactive({
    
    data <- raw_data()
    
    data$exp_date <- as.Date(data$exp_date)
    data$end_date <- as.Date(data$end_date) 
    data$year <- format(data$end_date, "%Y")
    # Pastikan jadi karakter
    #data$end_date <- ifelse(
    #grepl("^[0-9]+$", data$schedule),  # Jika angka (format Excel date)
    #format(as.Date(as.numeric(data$schedule), origin = "1899-12-30"), "%d %b"),  # Konversi ke format tanggal
    #ifelse(
    #grepl("^\\d{4}-\\d{2}-\\d{2}$", data$schedule),  # Jika format YYYY-MM-DD
    #format(as.Date(data$schedule), "%d %b"),  # Konversi ke format "22 Dec"
    #data$schedule  # Pertahankan nilai asli jika tidak sesuai)
    
    
    return(data)
  })
  
  
  ### HOME
  
  # Filter data based on type1 == "Certification"
  certification_data <- reactive({
    
    data <- raw_data()
    
    data$exp_date <- as.Date(data$exp_date)
    
    data %>% filter(type1 == "Certification")
  })
  
  # Function to count expired certifications based on the status
  count_expired <- function(status) {
    today <- Sys.Date()
    cert_data <- certification_data() %>% filter(!is.na(exp_date))
    
    if (status == "expired_this_year") {
      return(sum(year(cert_data$exp_date) == year(today) & cert_data$exp_date >= today))
    } else if (status == "expired_3_months") {
      return(sum(cert_data$exp_date > today & cert_data$exp_date <= today + months(3)))
    } else if (status == "expired_this_month") {
      return(sum(month(cert_data$exp_date) == month(today) & 
                   year(cert_data$exp_date) == year(today)))
    }
  }
  
  # Display count for the selected expired category
  output$expired_this_year <- renderText({
    count_expired("expired_this_year")
  })
  
  output$expired_3_months <- renderText({
    count_expired("expired_3_months")
  })
  
  output$expired_this_month <- renderText({
    count_expired("expired_this_month")
  })
  
  # Display certification details based on selected category
  observeEvent(input$btn_expired_this_year, {
    selected_data <- certification_data() %>%
      filter(year(exp_date) == year(Sys.Date()) & exp_date >= Sys.Date()) %>%
      mutate(No = row_number()) %>%
      select(No, `Employee Name` = name, `Department` = depart, 
             `Work Location` = work_loc, `Course Title` = title, 
             `Type` = type2, `Expired Date` = exp_date)
    
    selected_data$`Expired Date` <- format(selected_data$`Expired Date`, "%Y-%m-%d")
    
    output$cert_details <- renderTable({
      selected_data
    }, class = "table-custom")
    
    output$expired_title <- renderText({
      "Expired This Year"
    })
  })
  
  observeEvent(input$btn_expired_3_months, {
    selected_data <- certification_data() %>%
      filter(exp_date > Sys.Date() & exp_date <= Sys.Date() + months(3)) %>%
      mutate(No = row_number()) %>%
      select(No, `Employee Name` = name, `Department` = depart, 
             `Work Location` = work_loc, `Course Title` = title, 
             `Type` = type2, `Expired Date` = exp_date)
    
    selected_data$`Expired Date` <- format(selected_data$`Expired Date`, "%Y-%m-%d")
    
    output$cert_details <- renderTable({
      selected_data
    }, class = "table-custom")
    
    output$expired_title <- renderText({
      "Expired Next 3 Months"
    })
  })
  
  observeEvent(input$btn_expired_this_month, {
    selected_data <- certification_data() %>%
      filter(month(exp_date) == month(Sys.Date()) & year(exp_date) == year(Sys.Date())) %>%
      mutate(No = row_number()) %>%
      select(No, `Employee Name` = name, `Department` = depart, 
             `Work Location` = work_loc, `Course Title` = title, 
             `Type` = type2, `Expired Date` = exp_date)
    
    selected_data$`Expired Date` <- format(selected_data$`Expired Date`, "%Y-%m-%d")
    
    output$cert_details <- renderTable({
      selected_data
    }, class = "table-custom")
    
    output$expired_title <- renderText({
      "Expired This Month"
    })
  })
  
  # Untuk tampilan default
  filtered_expired_this_month <- reactive({
    emp_data <- certification_data()
    req(nrow(emp_data) > 0)
    
    # Filter data yang expired bulan ini
    filtered_data <- emp_data %>% 
      filter(month(exp_date) == month(Sys.Date()) & year(exp_date) == year(Sys.Date())) %>%
      mutate(No = row_number()) %>%
      select(No, `Employee Name` = name, `Department` = depart, 
             `Work Location` = work_loc, `Course Title` = title, 
             `Type` = type2, `Expired Date` = exp_date)
    
    # Mengubah format expired date
    filtered_data$`Expired Date` <- format(filtered_data$`Expired Date`, "%Y-%m-%d")
    
    return(filtered_data)
  })
  
  observe({
    output$cert_details <- renderTable({
      filtered_expired_this_month()
    }, class = "table-custom")
    
    output$expired_title <- renderText({
      "Expired This Month"
    })
  })
  
  ### DASHBOARD
  
  cleaned_data <- reactive({
    data <- rv$data
    print(names(rv$data))
    data <- data %>% select(-exp_date, -cost)
    data <- na.omit(data)
    data$start_date <- as.Date(data$start_date)
    data$end_date <- as.Date(data$end_date)
    data$year <- format(data$end_date, "%Y")  # Tambahkan kolom tahun
    
    return(data)
  })
  
  observe({
    req(cleaned_data())
    
    # Ambil daftar tahun unik dari data
    years <- unique(format(cleaned_data()$end_date, "%Y"))
    years <- sort(years, decreasing = TRUE)  # Urutkan dari terbaru ke terlama
    
    # Perbarui pilihan dropdown tahun
    updateSelectInput(
      session,
      inputId = "year_choice",
      choices = years
    )
  })
  
  # Filter data berdasarkan tahun pilihan user
  filtered_data <- reactive({
    data <- cleaned_data() %>%
      filter(year == input$year_choice)  # Bandingkan langsung dengan kolom tahun
    req(nrow(data) > 0, cancelOutput = TRUE)  # Hentikan proses jika kosong
    return(data)
  })
  
  
  # Boxes
  output$total_hours <- renderUI({
    total_hours <- sum(filtered_data()$m_hours, na.rm = TRUE)
    tags$div(
      class = "card",
      style = "width: 175px; height: 135px; background-color: #8fd744ff; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
      tags$div(
        style = "font-size: 35px; font-weight: bold; text-align: center;",
        total_hours
      ),
      tags$div(
        style = "font-size: 20px; margin-top: 10px;",
        HTML("Total<br>Man Hours")
      )
    )
  })
  
  output$total_days <- renderUI({
    total_days <- sum(filtered_data()$m_days, na.rm = TRUE) 
    tags$div(
      class = "card",
      style = "width: 175px; height: 135px; background-color: #2a7b8eff; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
      tags$div(
        style = "font-size: 35px; font-weight: bold; text-align: center;",
        total_days
      ),
      tags$div(
        style = "font-size: 20px; margin-top: 10px;",
        HTML("Total<br>Man Days")
      )
    )
  })
  
  output$total_learn_hours <- renderUI({
    total_learn_hours <- sum(filtered_data()$learn_hours, na.rm = TRUE)  # Hitung total learning hours
    tags$div(
      class = "card",
      style = "width: 175px; height: 135px; background-color: #414487ff; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
      tags$div(
        style = "font-size: 35px; font-weight: bold; text-align: center;",
        total_learn_hours 
      ),
      tags$div(
        style = "font-size: 20px; margin-top: 10px;",
        HTML("Total<br>Learning Hours")  # Subtitle
      )
    )
  })
  
  # Best Employee
  output$best_employee_hours <- renderUI({
    best_employee <- filtered_data() %>%
      group_by(name) %>%
      summarise(total_hours = sum(m_hours, na.rm = TRUE)) %>%
      arrange(desc(total_hours)) %>%
      slice_head(n = 1)
    # Hasilkan output HTML
    HTML(paste0(
      "<span class='highlight'>", best_employee$name, "</span><br>",
      "Total Man Hours: ", best_employee$total_hours
    ))
  })
  
  output$best_employee_titles <- renderUI({
    best_employee <- filtered_data() %>%
      group_by(name) %>%
      summarise(total_titles = n_distinct(title)) %>%
      arrange(desc(total_titles)) %>%
      slice_head(n = 1)
    # Hasilkan output HTML
    HTML(paste0(
      "<span class='highlight'>", best_employee$name, "</span><br>",
      "Many Trainings Attended: ", best_employee$total_titles
    ))
  })
  
  # Treemap by Department
  output$treemap <- renderPlot({
    # Pastikan tidak ada nilai 0
    treemap_data <- filtered_data() %>%
      count(depart, wt = m_hours) %>%
      filter(n > 0)  # Hapus jika ada nilai 0
    req(nrow(treemap_data) > 0, cancelOutput = TRUE)  # Cegah error jika kosong
    
    # Tentukan kategori yang akan muncul di legend dan di chart
    treemap_data$show_in_chart <- ifelse(treemap_data$n >= 10, treemap_data$depart, NA)  # Tampil di chart
    treemap_data$show_in_legend <- ifelse(treemap_data$n < 10, treemap_data$depart, NA)  # Tampil di legend
    
    # Perluas palet warna dengan viridis
    n_categories <- n_distinct(treemap_data$depart)  # Hitung jumlah kategori unik
    extended_palette <- viridis(n_categories, option = "viridis")  # Gunakan palet viridis
    
    ggplot(treemap_data, aes(area = n, label = depart, fill = depart)) +
      geom_treemap() +
      geom_treemap_text(
        fontface = "bold", 
        colour = "white", 
        place = "center", 
        show.legend = FALSE  # Tidak menampilkan kategori yang ada di chart pada legend
      ) +
      scale_fill_manual(values = extended_palette) +  # Gunakan palet warna viridis
      theme_minimal() +
      guides(
        fill = guide_legend(
          title = "Department", 
          labels = treemap_data$show_in_legend  # Menampilkan kategori yang tidak dicantumkan di chart di legend
        )
      ) +
      theme(legend.position = "right") 
  })
  
  output$chart_title <- renderText({
    paste("Man Hours per Month in", input$year_choice)  # Judul dinamis
  })
  
  # Visualisasi lainnya, seperti line chart, pie chart, bar chart, dan map
  output$man_hours_line_chart <- renderPlot({
    req(filtered_data())
    
    data_chart <- filtered_data() %>%
      filter(!is.na(start_date), !is.na(end_date), !is.na(m_hours)) %>%
      mutate(
        start_date = as.Date(start_date),
        end_date = as.Date(end_date)
      ) %>%
      filter(start_date <= end_date) %>%
      mutate(
        total_days = as.numeric(end_date - start_date) + 1,
        man_hours_per_day = m_hours / total_days
      ) %>%
      rowwise() %>%
      mutate(date_seq = list(seq.Date(start_date, end_date, by = "day"))) %>%
      unnest(date_seq) %>%
      mutate(
        month_num = month(date_seq),      # Ambil nomor bulan (1 = Januari, 2 = Februari, dst)
        month_label = month(date_seq, label = TRUE, abbr = FALSE, locale = "id_ID") # Nama bulan dalam Bahasa Indonesia
      ) %>%
      group_by(month_num, month_label) %>%
      summarise(total_man_hours = sum(man_hours_per_day, na.rm = TRUE), .groups = "drop") %>%
      arrange(month_num) %>%  # Pastikan urutan bulan benar sebelum cumsum()
      mutate(cumulative_man_hours = cumsum(total_man_hours)) %>%  # Hitung kumulatif
      ungroup()
    
    # Plot
    ggplot(data_chart, aes(x = factor(month_num, levels = 1:12), y = cumulative_man_hours, group = 1)) +
      geom_col(fill = "#fde725ff", alpha = 0.6) +  # Bar chart harus pakai cumulative_man_hours
      geom_line(color = "#414487ff", linewidth = 1) +  # Line chart tetap kumulatif
      geom_point(color = '#440154ff', size = 3) +
      geom_text(aes(label = round(cumulative_man_hours, 1)), vjust = -1.5, size = 3, color = 'navy', fontface = 'bold') +
      theme_minimal() +
      labs(
        title = "Cumulative Man Hours Per Month",
        x = "Bulan",
        y = "Man Hours"
      ) +
      theme(
        plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
        axis.text.x = element_text(angle = 45, hjust = 1),
        legend.position = "none",
        plot.margin = margin(20, 20, 20, 20)
      ) +
      scale_x_discrete(labels = unique(data_chart$month_label)) +  # Pakai nama bulan
      coord_cartesian(clip = 'off')
  })
  
  #PIEEE CHARTTTTTTTTTTT
  output$event_cat_pie <- renderPlotly({
    req(filtered_data())
    
    # Menghitung jumlah judul yang unik per kategori type1
    pie_data <- filtered_data() %>%
      group_by(event_cat, title) %>%  # Grup berdasarkan type1 dan title
      summarise(unique_titles = n_distinct(title), .groups = "drop") %>%  # Hitung jumlah judul unik per kategori
      count(event_cat, wt = unique_titles) %>%  # Hitung total jumlah berdasarkan kategori type1
      arrange(desc(n))  # Urutkan berdasarkan jumlah terbanyak
    
    # Hitung persentase
    pie_data <- pie_data %>%
      mutate(percentage = round(n / sum(n) * 100))  # Bulatkan persentase
    
    # Pilih palet warna dari viridis
    viridis_colors <- viridis(length(unique(pie_data$event_cat)))
    
    # Buat Pie Chart dengan Plotly
    plot_ly(
      pie_data, 
      labels = ~event_cat, 
      values = ~percentage, 
      type = "pie",
      textinfo = "percent",  # Menampilkan label dan persentase di dalam pie
      insidetextfont = list(color = "white", size = 10),  # Perbesar teks di dalam pie
      hoverinfo = "text",
      text = ~paste("Tipe:", event_cat, "<br>", percentage, "%"),  # Informasi tambahan pada hover
      marker = list(colors = viridis_colors)
    ) %>%
      layout(
        title = list(
          text = "<b>Event Category</b>", 
          x = 0.5, 
          font = list(size = 15)  # Perbesar ukuran judul
        ),
        margin = list(l = 5, r = 5, t = 25, b = 5),
        width = 235,
        height = 235,
        showlegend = TRUE,
        legend = list(
          font = list(size = 10),  # Perkecil ukuran font legend
          orientation = "h",  # Posisikan legend secara horizontal
          x = 0.5,  # Pusatkan legend secara horizontal
          xanchor = "center",  # Pusatkan anchor pada tengah
          y = -0.2  # Tempatkan legend di bawah chart
        )
      )
  })
  
  output$type1_pie <- renderPlotly({
    req(filtered_data())
    
    # Menghitung jumlah judul yang unik per kategori type1
    pie_data <- filtered_data() %>%
      group_by(type1, title) %>%  # Grup berdasarkan type1 dan title
      summarise(unique_titles = n_distinct(title), .groups = "drop") %>%  # Hitung jumlah judul unik per kategori
      count(type1, wt = unique_titles) %>%  # Hitung total jumlah berdasarkan kategori type1
      arrange(desc(n))  # Urutkan berdasarkan jumlah terbanyak
    
    # Hitung persentase
    pie_data <- pie_data %>%
      mutate(percentage = round(n / sum(n) * 100))  # Bulatkan persentase
    
    # Pilih palet warna dari viridis
    viridis_colors <- viridis(length(unique(pie_data$type1)))
    
    # Buat Pie Chart dengan Plotly
    plot_ly(
      pie_data, 
      labels = ~type1, 
      values = ~percentage, 
      type = "pie",
      textinfo = "percent",  # Menampilkan label dan persentase di dalam pie
      insidetextfont = list(color = "white", size = 10),  # Perbesar teks di dalam pie
      hoverinfo = "text",
      text = ~paste("Tipe:", type1, "<br>", percentage, "%"),  # Informasi tambahan pada hover
      marker = list(colors = viridis_colors)
    ) %>%
      layout(
        title = list(
          text = "<b>Type 1</b>", 
          x = 0.5, 
          font = list(size = 15)  # Perbesar ukuran judul
        ),
        margin = list(l = 5, r = 5, t = 25, b = 5),
        width = 235,
        height = 235,
        showlegend = TRUE,
        legend = list(
          font = list(size = 10),  # Perkecil ukuran font legend
          orientation = "h",  # Posisikan legend secara horizontal
          x = 0.5,  # Pusatkan legend secara horizontal
          xanchor = "center",  # Pusatkan anchor pada tengah
          y = -0.2  # Tempatkan legend di bawah chart
        )
      )
  })
  
  output$type2_pie <- renderPlotly({
    req(filtered_data())
    
    # Menghitung jumlah judul yang unik per kategori type1
    pie_data <- filtered_data() %>%
      group_by(type2, title) %>%  # Grup berdasarkan type1 dan title
      summarise(unique_titles = n_distinct(title), .groups = "drop") %>%  # Hitung jumlah judul unik per kategori
      count(type2, wt = unique_titles) %>%  # Hitung total jumlah berdasarkan kategori type1
      arrange(desc(n))  # Urutkan berdasarkan jumlah terbanyak
    
    # Hitung persentase
    pie_data <- pie_data %>%
      mutate(percentage = round(n / sum(n) * 100))  # Bulatkan persentase
    
    # Pilih palet warna dari viridis
    viridis_colors <- viridis(length(unique(pie_data$type2)))
    
    # Buat Pie Chart dengan Plotly
    plot_ly(
      pie_data, 
      labels = ~type2, 
      values = ~percentage, 
      type = "pie",
      textinfo = "percent",  # Menampilkan label dan persentase di dalam pie
      insidetextfont = list(color = "white", size = 10),  # Perbesar teks di dalam pie
      hoverinfo = "text",
      text = ~paste("Tipe:", type2, "<br>", percentage, "%"),  # Informasi tambahan pada hover
      marker = list(colors = viridis_colors)
    ) %>%
      layout(
        title = list(
          text = "<b>Type 2</b>", 
          x = 0.5, 
          font = list(size = 15)  # Perbesar ukuran judul
        ),
        margin = list(l = 5, r = 5, t = 25, b = 5),
        width = 235,
        height = 235,
        showlegend = TRUE,
        legend = list(
          font = list(size = 10),  # Perkecil ukuran font legend
          orientation = "h",  # Posisikan legend secara horizontal
          x = 0.5,  # Pusatkan legend secara horizontal
          xanchor = "center",  # Pusatkan anchor pada tengah
          y = -0.2  # Tempatkan legend di bawah chart
        )
      )
  })
  
  output$top_venues_bar <- renderPlot({
    # Prepare the data: Count unique venue-title combinations
    venue_data <- filtered_data() %>%
      group_by(venue, title) %>%
      summarise(unique_titles = n_distinct(title), .groups = "drop") %>%
      count(venue, wt = unique_titles) %>%  # Count unique venue-title combinations
      arrange(desc(n))
    
    top_venues <- venue_data %>%
      slice_head(n = 5) 
    
    # Create the horizontal bar chart with viridis colors
    ggplot(top_venues, aes(x = reorder(venue, n), y = n, fill = venue)) +
      geom_bar(stat = "identity", show.legend = FALSE) +  # Horizontal bars
      geom_text(aes(label = n), hjust = 1.5, size = 6, colour = "white", fontface = "bold") +  # Adjust labels
      coord_flip() +  # Flip to horizontal bars
      scale_fill_viridis_d() +  # Use viridis color palette for bars
      theme_minimal() +  # Minimal theme for clean background
      theme(
        panel.grid = element_blank(),        # Remove grid lines
        axis.title.y = element_blank(),      # Remove y-axis title
        axis.title.x = element_blank(),      # Remove x-axis title
        axis.text = element_text(size = 12), # Adjust text size for clarity
        axis.text.y = element_text(hjust = 1) # Align y-axis labels properly
      )
  })
  
  ### RECORD
  
  # Search Nama
  observe({
    updateSelectizeInput(
      session,
      inputId = "search_name",
      choices = unique(processed_data()$name),  # Daftar nama unik dari data
      server = TRUE   # Mode server untuk menangani data besar
    )
  })
  
  # Filter data berdasarkan nama yang dipilih
  selected_employee <- reactive({
    req(input$search_name)
    
    data <- processed_data() %>% filter(name == input$search_name)
    
    if (input$filter_year != "All") {
      data <- data %>% filter(year == as.numeric(input$filter_year))  # Filter by year
    }
    
    return(data)
    
  })
  
  observe({
    req(processed_data())
    
    # Ambil daftar tahun unik dari data
    years <- unique(processed_data()$year)
    years <- sort(years, decreasing = TRUE)  # Urutkan dari terbaru ke terlama
    
    # Perbarui pilihan dropdown tahun
    updateSelectInput(
      session,
      inputId = "filter_year",
      choices = c("All",years),
      selected = "All"  
    )
  })
  
  filtered_employee <- reactive({
    emp_data <- selected_employee()
    req(nrow(emp_data) > 0)
    
    # Periksa apakah tombol "Specific Year" dipilih
    if (input$filter_year == "All") {
      emp_data <- emp_data %>% filter(year == input$filter_year)
    }
    
    return(emp_data)
  })
  
  # Baris Pertama: Detail Karyawan
  output$employee_details <- renderUI({
    emp_data <- selected_employee()
    req(nrow(emp_data) > 0)
    
    tagList(
      div(
        class = "box-custom",
        style = "margin-top: 0px; justify-content: center",
        h3(style = "font-size: 40px; color: #333; font-weight: bold; margin-top: 0px", emp_data$name[1]),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Department:</strong> {emp_data$depart[1]}"))),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Status:</strong> {emp_data$emp_status[1]}"))),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Work Location:</strong> {emp_data$work_loc[1]}")))
      )
    )
  })
  
  output$training_stats <- renderUI({
    emp_data <- selected_employee()
    req(nrow(emp_data) > 0)
    
    total_course <- length(unique(emp_data$title))
    total_days <- sum(emp_data$m_days, na.rm = TRUE)
    total_hours <- sum(emp_data$m_hours, na.rm = TRUE)
    total_learn_hours <- sum(emp_data$learn_hours, na.rm = TRUE)
    total_participants <- nrow(emp_data)  # Misalnya jumlah peserta
    
    tags$div(
      style = "display: flex; gap: 20px; justify-content: space-evenly; margin-top: 20px; flex-wrap: wrap;",
      
      # Square cards for Total Courses
      tags$div(
        class = "card",
        style = "width: 120px; height: 150px; background-color: #0093ad; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
        tags$div(
          style = "font-size: 45px; font-weight: bold;",
          total_course
        ),
        tags$div(
          style = "font-size: 16px; margin-top: 10px;",
          HTML("Total<br>Courses")
        )
      ),
      
      # Square cards for Total Days
      tags$div(
        class = "card",
        style = "width: 120px; height: 150px; background-color: #006b89; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
        tags$div(
          style = "font-size: 45px; font-weight: bold;",
          total_days
        ),
        tags$div(
          style = "font-size: 16px; margin-top: 10px;",
          HTML("Total<br>Days")
        )
      ),
      
      # Square cards for Total Hours
      tags$div(
        class = "card",
        style = "width: 120px; height: 150px; background-color: #134a6e; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
        tags$div(
          style = "font-size: 45px; font-weight: bold;",
          total_hours
        ),
        tags$div(
          style = "font-size: 16px; margin-top: 10px;",
          HTML("Total<br>Hours")
        )
      ),
      
      # Square cards for Learning Hours
      tags$div(
        class = "card",
        style = "width: 120px; height: 150px; background-color: #1f2855; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
        tags$div(
          style = "font-size: 45px; font-weight: bold;",
          total_learn_hours
        ),
        tags$div(
          style = "font-size: 16px; margin-top: 10px;",
          HTML("Learning<br>Hours")
        )
      )
    )
  })
  
  output$training_history <- renderUI({
    emp_data <- selected_employee()
    req(nrow(emp_data) > 0)
    
    # Urutkan data berdasarkan tahun terbaru
    emp_data <- emp_data %>% arrange(desc(year))
    
    # Pisahkan berdasarkan tahun
    years <- unique(emp_data$year)
    history <- lapply(years, function(yr) {
      year_data <- emp_data %>% filter(year == yr)
      
      tags$div(
        h3(style = "margin-top: 20px; text-transform: uppercase; color: #134a6e; font-size: 40px; font-weight: bold; margin-top: 0;", yr),
        # Gunakan ul untuk membuat daftar pelatihan dengan bullet point
        tags$ul(
          style = "list-style: disc; padding-left: 30px;",  # Menggunakan bullet point dan memberikan indentasi
          lapply(1:nrow(year_data), function(i) {
            tags$li(
              style = "margin-bottom: 20px;",
              # Judul pelatihan dan detail lainnya
              tags$span(
                style = "font-weight: bold; font-size: 20px;",
                year_data$title[i]
              ),
              tags$p(
                style = "font-size: 16px;",
                glue::glue("{year_data$venue[i]}, {year_data$start_date[i]} s.d. {year_data$end_date[i]}")
              ),
              # Label kategori
              tags$div(
                style = "display: flex; gap: 10px;",
                tags$div(style = "padding: 3px 6px; background-color: #2c5f78; color: white; border-radius: 5px;", year_data$type1[i]),
                tags$div(style = "padding: 3px 6px; background-color: #3e849e; color: white; border-radius: 5px;", year_data$event_cat[i]),
                tags$div(style = "padding: 3px 6px; background-color: #51acc5; color: white; border-radius: 5px;", year_data$type2[i])
              )
            )
          })
        )
      )
    })
    
    # Menggabungkan semua history pelatihan berdasarkan tahun
    do.call(tagList, history)
  })
  
   # Reactive untuk menyimpan data training history
  export_data <- reactive({
    emp_data <- selected_employee()
    req(nrow(emp_data) > 0)
    
    df <- emp_data %>% select(year, title, venue, start_date, type1, event_cat, type2)
    colnames(df) <- c("Year", "Course Title", "Venue", "Start Date", "Type 1", "Event Category", "Type 2")
    
    emp_info <- list(
      Name = emp_data$name[1],
      Department = emp_data$depart[1],
      Status = emp_data$emp_status[1],
      Work_Location = emp_data$work_loc[1]
    )
    
    return(list(info = emp_info, data = df))
  })
  
  generate_word_report <- function(emp_info, df) {
    doc <- read_docx()
    
    landscape_section <- block_section(
      prop_section(
        page_size(width = 11, height = 8.5, orient = "landscape")
      )
    )
    
    doc <- doc %>%
      body_add_par("Employee Training Report", style = "centered") %>%
      body_add_par(" ") %>%
      body_add_par(paste("Name:", emp_info$Name)) %>%
      body_add_par(paste("Department:", emp_info$Department)) %>%
      body_add_par(paste("Status:", emp_info$Status)) %>%
      body_add_par(paste("Work Location:", emp_info$Work_Location)) %>%
      body_add_par(" ")
    
    flextable_df <- flextable(df) %>%
      theme_vanilla() %>%
      autofit() %>%
      set_table_properties(layout = "autofit")
    
    doc <- doc %>%
      body_add_flextable(flextable_df) %>%
      body_add_par(" ") %>%
      body_end_block_section(landscape_section)
    
    file_path <- tempfile(fileext = ".docx")
    print(doc, target = file_path)
    return(file_path)
  }
  
  output$export_button <- downloadHandler(
    filename = function() {
      emp_data <- export_data()  
      emp_name <- emp_data$info$Name  
      paste("TrainingRecord_", emp_name, "_", Sys.Date(), ".docx", sep = "") 
    },
    content = function(file) {
      emp_data <- export_data()
      file_path <- generate_word_report(emp_data$info, emp_data$data)
      file.copy(file_path, file)
    }
  )
  
  
  ### COURSE
  # Search Nama
  observe({
    updateSelectizeInput(
      session,
      inputId = "search_title",
      choices = unique(processed_data()$title),
      server = TRUE 
    )
  })
  
  # Filter data berdasarkan judul yang dipilih
  selected_title <- reactive({
    req(input$search_title,input$title_year)
    
    data <- processed_data() %>% filter(title == input$search_title, year == as.numeric(input$title_year))
    
    return(data)
    
  })
  
  observe({
    req(processed_data())
    
    years <- unique(processed_data()$year)
    years <- sort(years, decreasing = TRUE)  # Urutkan dari terbaru ke terlama
    
    updateSelectInput(
      session,
      inputId = "title_year",
      choices = years
    )
  })
  
  title_data <- reactive({
    selected_title()
  })
  
  output$course_details <- renderUI({
    title_data <- selected_title()
    req(nrow(title_data) > 0)
    title_data$start_date = format(as.Date(title_data$start_date), "%d-%m-%Y")
    title_data$end_date = format(as.Date(title_data$end_date), "%d-%m-%Y")
    
    tagList(
      div(
        class = "box-custom",
        style = "margin-top: 0px; justify-content: center",
        h3(style = "font-size: 40px; color: #333; font-weight: bold; margin-top: 0px", title_data$title[1]),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Start Date:</strong> {title_data$start_date[1]}"))),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>End Date:</strong> {title_data$end_date[1]}"))),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Location:</strong> {title_data$venue[1]}"))),
        p(style = "font-size: 18px; color: #555;", HTML(glue::glue("<strong>Vendor:</strong> {title_data$vendor[1]}"))),
        tags$div(
          style = "display: flex; gap: 10px;",
          tags$div(style = "padding: 3px 6px; background-color: #2c5f78; color: white; border-radius: 5px;", title_data$type1[1]),
          tags$div(style = "padding: 3px 6px; background-color: #3e849e; color: white; border-radius: 5px;", title_data$event_cat[1]),
          tags$div(style = "padding: 3px 6px; background-color: #51acc5; color: white; border-radius: 5px;", title_data$type2[1])
        )
      )
    )
  })
  
  output$course_stats <- renderUI({
    title_data <- selected_title()
    req(nrow(title_data) > 0)
    
    total_participants <- length(unique(title_data$name))
    total_days <- title_data$m_days[1]
    total_hours <- title_data$m_hours[1]
    total_learn_hours <- title_data$learn_hours[1]
    
    tags$div(
      style = "margin-top: 20px; display: flex; flex-direction: column; gap: 20px;",
      
      # Baris pertama (Total Participants dan Total Days)
      tags$div(
        style = "display: flex; gap: 20px; justify-content: center;",
        tags$div(
          class = "card",
          style = "width: 120px; height: 150px; background-color: #0093ad; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
          tags$div(
            style = "font-size: 45px; font-weight: bold;",
            total_participants
          ),
          tags$div(
            style = "font-size: 16px; margin-top: 10px;",
            HTML("Total<br>Participants")
          )
        ),
        tags$div(
          class = "card",
          style = "width: 120px; height: 150px; background-color: #006b89; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
          tags$div(
            style = "font-size: 45px; font-weight: bold;",
            total_days
          ),
          tags$div(
            style = "font-size: 16px; margin-top: 10px;",
            HTML("Total<br>Days")
          )
        )
      ),
      
      # Baris kedua (Man Hours dan Learning Hours)
      tags$div(
        style = "display: flex; gap: 20px; justify-content: center;",
        tags$div(
          class = "card",
          style = "width: 120px; height: 150px; background-color: #134a6e; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
          tags$div(
            style = "font-size: 45px; font-weight: bold;",
            total_hours
          ),
          tags$div(
            style = "font-size: 16px; margin-top: 10px;",
            HTML("Total<br>Hours")
          )
        ),
        tags$div(
          class = "card",
          style = "width: 120px; height: 150px; background-color: #1f2855; color: white; font-size: 20px; font-weight: bold; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 15px; text-align: center; border-radius: 10px;",
          tags$div(
            style = "font-size: 45px; font-weight: bold;",
            total_learn_hours
          ),
          tags$div(
            style = "font-size: 16px; margin-top: 10px;",
            HTML("Learning<br>Hours")
          )
        )
      )
    )
    
  })
  
  
  output$course_participants <- renderTable({
    title_data <- selected_title()
    req(nrow(title_data) > 0)
    
    filtered_data <- title_data %>%
      select(name, depart, emp_status, work_loc) %>%
      arrange(name) %>%
      mutate(No = row_number()) %>%
      select(No, name, depart, emp_status, work_loc) %>%
      rename(
        "No" = No,
        "Name" = name,
        "Department" = depart,
        "Employee Status" = emp_status,
        "Work Location" = work_loc
      )
    
    filtered_data
  }, class = "table-custom")
  
  observe({
    req(processed_data())
    years <- unique(processed_data()$year)
    years <- sort(years, decreasing = TRUE)
    updateSelectInput(
      session,
      inputId = "title_year",
      choices = years
    )
  })
  
  selected_title <- reactive({
    req(input$search_title, input$title_year)
    data <- processed_data() %>% filter(title == input$search_title, year == as.numeric(input$title_year))
    return(data)  
  })
  
  export_title <- reactive({
    title_data <- selected_title()
    req(nrow(title_data) > 0)
    
    df <- title_data %>%
      select(name, depart, emp_status, work_loc) %>%
      arrange(name) %>%
      mutate(No = row_number()) %>%
      select(No, name, depart, emp_status, work_loc) %>%
      rename(
        "No" = No,
        "Name" = name,
        "Department" = depart,
        "Employee Status" = emp_status,
        "Work Location" = work_loc
      )
    
    course_info <- list(
      Title = title_data$title[1],
      Start_Date = format(as.Date(title_data$start_date[1]), "%d-%m-%Y"),
      End_Date = format(as.Date(title_data$end_date[1]), "%d-%m-%Y"),
      Location = title_data$venue[1],
      Vendor = title_data$vendor[1]
    )
    
    return(list(info = course_info, data = df))
  })
  
  generate_title_report <- function(course_info, df) {
    doc <- read_docx()
    
    doc <- doc %>%
      body_add_par("Course Training Report", style = "centered") %>%
      body_add_par(" ") %>%
      body_add_par(paste("Title:", course_info$Title)) %>%
      body_add_par(paste("Start Date:", course_info$Start_Date)) %>%
      body_add_par(paste("End Date:", course_info$End_Date)) %>%
      body_add_par(paste("Location:", course_info$Location)) %>%
      body_add_par(paste("Vendor:", course_info$Vendor)) %>%
      body_add_par(" ")
    
    flextable_df <- flextable(df) %>%
      theme_vanilla() %>%
      autofit()
    
    doc <- doc %>%
      body_add_flextable(flextable_df) %>%
      body_add_par(" ")
    
    file_path <- tempfile(fileext = ".docx")
    print(doc, target = file_path)
    return(file_path)
  }
  
  output$export_participants <- downloadHandler(
    filename = function() {
      year = input$title_year
      title_data <- export_title()  
      train_title <- title_data$info$Title
      paste("TrainingDetails_", train_title,"_" ,year, "_", Sys.Date(), ".docx", sep = "") 
    },
    content = function(file) {
      course_data <- export_title()
      file_path <- generate_title_report(course_data$info, course_data$data)
      file.copy(file_path, file)
    }
  )
  
  ## DATA
  # Reactively load the uploaded Excel data
  observeEvent(input$file_upload, {
    req(input$file_upload)  # Pastikan file sudah di-upload
    
    uploaded_data <- read_excel(input$file_upload$datapath)
    
    # Logika untuk validasi data Excel (memeriksa kolom yang diperlukan)
    required_columns <- c("name", "depart", "emp_status", "work_loc", "cost", "title", 
                          "start_date", "end_date", "m_days", "m_hours", "learn_hours", 
                          "venue", "vendor", "type1", "event_cat", "type2")
    
    missing_columns <- setdiff(required_columns, colnames(uploaded_data))
    
    if (length(missing_columns) > 0) {
      showNotification(paste("Missing required columns in the uploaded file:", 
                             paste(missing_columns, collapse = ", ")), type = "error")
      return()
    }
    
    # Hanya memeriksa baris yang tidak kosong
    non_empty_rows <- uploaded_data[complete.cases(uploaded_data), ]
    
    # Memeriksa apakah ada kolom yang kosong pada baris yang memiliki data
    missing_values <- non_empty_rows %>%
      filter(if_any(everything(), ~ is.na(.)))  # Memeriksa jika ada kolom yang kosong pada baris yang terisi
    
    if (nrow(missing_values) > 0) {
      showNotification("Some required fields are missing in the uploaded data.", type = "error")
      return()
    }
    
    # Validasi m_days, m_hours, learn_hours untuk memastikan tidak ada nilai negatif atau kosong
    invalid_numbers <- uploaded_data %>%
      filter(m_days < 0 | m_hours < 0 | learn_hours < 0)
    
    if (nrow(invalid_numbers) > 0) {
      showNotification("Man Days, Man Hours, and Learning Hours cannot be negative.", type = "error")
      return()
    }
    
    # Gabungkan data yang di-upload dengan data yang sudah ada
    rv$data <- rbind(rv$data, uploaded_data)
    
    # Simpan data gabungan ke file Excel
    write_xlsx(rv$data, "raw_data_dummy.xlsx")
    
    showNotification("Data from the uploaded file has been successfully added to the existing data.", type = "message")
    
    # Tampilkan data yang digabungkan di preview_new_data
    output$preview_new_data <- renderDT({
      df <- uploaded_data  # Ambil data gabungan dari rv$data
      colnames(df) <- c(
        "Name", "Department", "Employee Status", "Work Location", "Cost Center", 
        "Course Title", "Start Training Date", "End Training Date", "Man Days", "Man Hours", "Learning Hours",
        "Venue", "Vendor", "Type 1", "Event Category", "Type 2", "Expired Date"
      )
      datatable(df, options = list(dom = "t", scrollX=TRUE), rownames = FALSE)
    })
  })
  
  # Menambahkan data baru dengan validasi
  observeEvent(input$add_data, {
    
    # Buat data frame baru dengan tibble
    new_entry <- tibble(
      name = input$name,
      depart = input$depart,
      emp_status = input$emp_status,
      work_loc = input$work_loc,
      cost = input$cost,
      title = input$title,
      start_date = input$start_date,
      end_date = input$end_date,
      m_days = as.numeric(input$m_days),
      m_hours = as.numeric(input$m_hours),
      learn_hours = as.numeric(input$learn_hours),
      venue = input$venue,
      vendor = input$vendor,
      type1 = input$type1,
      event_cat = input$event_cat,
      type2 = input$type2,
      exp_date = input$exp_date
    )
    
    # Validasi apakah ada kolom kosong (selain exp_date)
    required_cols <- c("name", "depart", "emp_status", "work_loc", "cost", "title",
                       "start_date", "end_date", "m_days", "m_hours", "learn_hours",
                       "venue", "vendor", "type1", "event_cat", "type2")
    
    # Cek apakah ada nilai kosong
    missing_values <- new_entry %>%
      select(all_of(required_cols)) %>%
      summarise_all(~ any(is.na(.) | . == "")) %>%
      unlist()
    
    # Gantilah NA dengan FALSE agar dapat dievaluasi
    missing_values[is.na(missing_values)] <- FALSE
    
    # Cek apakah ada kolom yang kosong
    if (any(missing_values)) {
      showNotification("Error: Semua kolom wajib diisi, kecuali exp_date.", type = "error")
      return()
    }
    
    
    # Jika tidak ada missing values, lanjutkan menambah data
    rv$data <- if (is.null(rv$data) || nrow(rv$data) == 0) new_entry else bind_rows(rv$data, new_entry)
    
    # Tampilkan data baru
    output$preview_new_data <- renderDT({
      datatable(new_entry, options = list(dom = "t", scrollX = TRUE), rownames = FALSE)
    })
    
    # Notifikasi sukses
    showNotification("Data successfully added!", type = "message")
  })
  
  
  # View and Edit table
  #Edit table
  # Render DataTable untuk edit data
  output$edit_table <- renderDT({
    
    df <- rv$data
    
    # Pastikan format tanggal tetap sebagai Date
    df <- df %>%
      mutate(
        delete = paste0('<button class="btn-delete" onclick="Shiny.setInputValue(\'delete_row\', ', row_number(), ', {priority: \'event\'})">Delete</button>'),
        start_date = as.Date(start_date, origin = "1970-01-01"),
        end_date = as.Date(end_date, origin = "1970-01-01"),
        exp_date = as.Date(exp_date, origin = "1970-01-01")
      )
    
    colnames(df) <- c(
      "Name", "Department", "Employee Status", "Work Location", "Cost Center", 
      "Course Title", "Start Training Date", "End Training Date", "Man Days", 
      "Man Hours", "Learning Hours", "Venue", "Vendor", "Type 1", 
      "Event Category", "Type 2", "Expired Date", "Delete"
    )
    
    # Tampilkan DataTable dengan kolom tanggal editable menggunakan DatePicker
    datatable(
      df,
      escape = FALSE, # Prevent escaping HTML
      editable = TRUE,
      options = list(pageLength = 5, scrollX = TRUE),
      rownames = FALSE
    ) %>%
      formatDate(columns = c("Start Training Date", "End Training Date", "Expired Date"), method = "toLocaleDateString")
  })
  
  
  # Event Listener untuk Menyimpan Perubahan
  library(openxlsx)
  
  observeEvent(input$edit_table_cell_edit, {
    info <- input$edit_table_cell_edit
    
    # Ambil data reaktif
    df <- rv$data
    
    # Dapatkan nama kolom yang diedit
    col_name <- colnames(df)[info$col + 1]  # +1 karena DataTable tidak menghitung rownames
    
    # Periksa apakah kolom yang diedit adalah kolom tanggal
    if (col_name %in% c("Start Training Date", "End Training Date", "Expired Date")) {
      new_date <- tryCatch(
        as.Date(info$value, format="%Y-%m-%d"),
        error = function(e) NA  # Jika gagal konversi, masukkan NA
      )
      
      if (!is.na(new_date)) {
        rv$data[[col_name]][info$row] <- new_date
      } else {
        showNotification("Invalid date format. Please enter a valid date.", type = "error")
      }
    } else {
      # Jika bukan kolom tanggal, langsung update nilai
      rv$data[[col_name]][info$row] <- as.character(info$value)  # Pastikan kompatibilitas tipe data
    }
    
  })
  
  observeEvent(input$save_changes, {
    file_path <- "raw_data_dummy.xlsx"
    
    tryCatch({
      write.xlsx(rv$data, file = file_path, overwrite = TRUE)
      showNotification("Changes saved successfully!", type = "message")
    }, error = function(e) {
      showNotification("Error saving file: Please check permissions.", type = "error")
    })
  })
  
  
  
  # Save edits
  # Observe delete button clicks by matching the button ID
  observeEvent(input$delete_row, {
    row_to_delete <- as.numeric(input$delete_row)
    
    if (!is.na(row_to_delete) && row_to_delete > 0 && row_to_delete <= nrow(rv$data)) {
      rv$data <- rv$data[-row_to_delete, ]  # Hapus baris dari data
      write_xlsx(rv$data, "raw_data_dummy.xlsx")  # Simpan perubahan ke file
      showNotification("Row deleted successfully and changes saved to Excel file!", type = "message")
    }
  })
  
  # Preview of new data
  output$preview_new_data <- renderDT({
    datatable(data.frame(
      name = input$name,
      depart = input$depart,
      emp_status = input$emp_status,
      work_loc = input$work_loc,
      cost = input$cost,
      title = input$title,
      start_date = input$start_date,
      end_date = input$end_date,
      m_days = input$m_days,
      m_hours = input$m_hours,
      learn_hours = input$learn_hours,
      venue = input$venue,
      vendor = input$vendor,
      type1 = input$type1,
      event_cat = input$event_cat,
      type2 = input$type2,
      exp_date = ifelse(input$exp_date == "", NA, input$exp_date)
    ), options = list(dom = "t"), rownames = FALSE)
  })
  
  # Output Tabel Ekspor
  output$data_table <- renderDT({
    df <- rv$data
    
    colnames(df) <- c(
      "Name", "Department", "Employee Status", "Work Location", "Cost Center", 
      "Course Title", "Start Training Date", "End Training Date", "Man Days", "Man Hours", "Learning Hours",
      "Venue", "Vendor", "Type 1", "Event Category", "Type 2", "Expired Date"
    )
    
    datatable(df, 
              extensions = 'Buttons', 
              options = list(
                dom = 'Bfrtip', 
                pageLength = -1,
                buttons = list(
                  list(
                    extend = 'excel', 
                    text = 'Export Excel', 
                    exportOptions = list(
                      modifier = list(
                        page = 'all'
                      )
                    ),
                    # Menambahkan kelas dan style kustom untuk tombol
                    className = 'action-btn'
                  ),
                  list(
                    extend = 'csv',
                    text = 'Export CSV',
                    exportOptions = list(
                      modifier = list(
                        page = 'all'
                      )
                    ),
                    className = 'action-btn'
                  )
                ),
                pageLength = 10,
                scrollX = TRUE
              )
    )
  }, server = TRUE)
  
  output$pgn <- renderImage({
    list(src="www/LOGOPGN.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn1 <- renderImage({
    list(src="www/LOGOPGN1.png",width= "300", height= "100")
  },deleteFile = F)
  
  output$pgn2 <- renderImage({
    list(src="www/LOGOPGN2.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn3 <- renderImage({
    list(src="www/LOGOPGN3.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn4 <- renderImage({
    list(src="www/LOGOPGN4.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn5 <- renderImage({
    list(src="www/LOGOPGN5.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn6 <- renderImage({
    list(src="www/LOGOPGN6.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn7 <- renderImage({
    list(src="www/LOGOPGN7.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$pgn8 <- renderImage({
    list(src="www/LOGOPGN8.png",width = "300", height = "100")
  },deleteFile = F)
  
  output$saka <- renderImage({
    list(src="www/saka1.jpg",width = "100%", height = "100%")
  },deleteFile = F)
  
  output$calista <- renderImage({
    list(src="www/calista.jpg",width = "240", height = "320")
  },deleteFile = F)
  
  output$rosita <- renderImage({
    list(src="www/rosita.jpg",width = "240", height = "320")
  },deleteFile = F)
  
  # Reactive values for status message and color
  upload_message <- reactiveVal("File belum diunggah.")
  upload_color <- reactiveVal("#0056a1")  # Default color: blue
  
  # Feedback Data
  feedback_data <- reactive({
    req(input$feedback_file)
    file <- input$feedback_file$datapath
    tryCatch({
      # Membaca file dan mengubah nama kolom
      df <- read_excel(file)
      df <- df %>% 
        setNames(c("Timestamp", "Judul Pelatihan", "Tempat Pelaksanaan", "Nama Peserta",
                   "Departemen", "Kejelasan Tujuan", "Ketercapaian Tujuan", "Kesesuaian Materi",
                   "Penerapan Materi", "Tambahan Pengetahuan", "Tingkat Kesulitan", 
                   "Kemudahan Pemahaman", "Fasilitator Menguasai", "Kompetensi Fasilitator", 
                   "Keaktifan Fasilitator", "Penilaian Fasilitator", "Total Penilaian", 
                   "Point Utama Materi", "Kritik dan Saran"))
      
      # Pastikan kolom yang dibutuhkan ada di file
      if (!all(c("Timestamp", "Judul Pelatihan", "Tempat Pelaksanaan", "Nama Peserta",
                 "Departemen", "Kejelasan Tujuan", "Ketercapaian Tujuan", "Kesesuaian Materi",
                 "Penerapan Materi", "Tambahan Pengetahuan", "Tingkat Kesulitan", 
                 "Kemudahan Pemahaman", "Fasilitator Menguasai", "Kompetensi Fasilitator", 
                 "Keaktifan Fasilitator", "Penilaian Fasilitator", "Total Penilaian", 
                 "Point Utama Materi", "Kritik dan Saran") %in% colnames(df))) {
        showNotification("Error: File tidak sesuai format", type = "error")
        return(NULL)
      }
      
      # Convert columns to integer where needed
      df$`Tingkat Kesulitan` <- as.integer(str_extract(df$`Tingkat Kesulitan`, "^\\d+"))
      df$`Penilaian Fasilitator` <- as.integer(str_extract(df$`Penilaian Fasilitator`, "^\\d+"))
      df$`Total Penilaian` <- as.integer(str_extract(df$`Total Penilaian`, "^\\d+"))
      
      return(df)
    }, error = function(e) {
      showNotification("Error reading Feedback file", type = "error")
      NULL
    })
  })
  
  # Pre-Test Data
  pretest_data <- reactive({
    req(input$pretest_file)
    file <- input$pretest_file$datapath
    tryCatch({
      # Membaca file dan mengubah nama kolom
      df <- read_excel(file)
      colnames(df) <- c("Nama Peserta", "Departemen","Nilai Pre-Test")
      df %>% arrange(`Nama Peserta`)
    }, error = function(e) {
      showNotification("Error reading Pre-Test file", type = "error")
      NULL
    })
  })
  
  # Post-Test Data
  posttest_data <- reactive({
    req(input$posttest_file)
    file <- input$posttest_file$datapath
    tryCatch({
      # Membaca file dan mengubah nama kolom
      df <- read_excel(file)
      colnames(df) <- c("Nama Peserta", "Departemen", "Nilai Post-Test")
      df %>% arrange(`Nama Peserta`)
    }, error = function(e) {
      showNotification("Error reading Post-Test file", type = "error")
      NULL
    })
  })
  
  # Observer untuk error handling saat upload
  observe({
    tryCatch({
      # Feedback file validation
      if (!is.null(input$feedback_file)) {
        df <- readxl::read_excel(input$feedback_file$datapath)
        
        # Pastikan jumlah kolom yang ada sesuai dengan format (harus 19 kolom)
        if (ncol(df) != 19) {
          upload_message("Error: File Feedback harus memiliki 19 kolom.")
          upload_color("#b22222")  # Error: red
          return(NULL)
        }
        
        df <- df %>% 
          setNames(c("Timestamp", "Judul Pelatihan", "Tempat Pelaksanaan", "Nama Peserta",
                     "Departemen", "Kejelasan Tujuan", "Ketercapaian Tujuan", "Kesesuaian Materi",
                     "Penerapan Materi", "Tambahan Pengetahuan", "Tingkat Kesulitan", 
                     "Kemudahan Pemahaman", "Fasilitator Menguasai", "Kompetensi Fasilitator", 
                     "Keaktifan Fasilitator", "Penilaian Fasilitator", "Total Penilaian", 
                     "Point Utama Materi", "Kritik dan Saran"))
        
        if (!all(c("Timestamp", "Judul Pelatihan", "Tempat Pelaksanaan", "Nama Peserta",
                   "Departemen", "Kejelasan Tujuan", "Ketercapaian Tujuan", "Kesesuaian Materi",
                   "Penerapan Materi", "Tambahan Pengetahuan", "Tingkat Kesulitan", 
                   "Kemudahan Pemahaman", "Fasilitator Menguasai", "Kompetensi Fasilitator", 
                   "Keaktifan Fasilitator", "Penilaian Fasilitator", "Total Penilaian", 
                   "Point Utama Materi", "Kritik dan Saran") %in% colnames(df))) {
          upload_message("Error: File Feedback tidak sesuai format.")
          upload_color("#b22222")  # Error: red
        }
      }
      
      # Pre-Test file validation
      if (!is.null(input$pretest_file)) {
        df <- readxl::read_excel(input$pretest_file$datapath)
        
        # Pastikan jumlah kolom yang ada adalah 3 (Nama Peserta, Departemen, Nilai Pre-Test)
        if (ncol(df) != 3) {
          upload_message("Error: File Pre-Test harus memiliki 3 kolom.")
          upload_color("#b22222")  # Error: red
          return(NULL)
        }
        
        # Rename kolom sesuai dengan yang diinginkan
        colnames(df) <- c("Nama Peserta", "Departemen", "Nilai Pre-Test")
        
        if (!all(c("Nama Peserta", "Departemen", "Nilai Pre-Test") %in% colnames(df))) {
          upload_message("Error: File Pre-Test tidak sesuai format.")
          upload_color("#b22222")  # Error: red
        }
      }
      
      # Post-Test file validation
      if (!is.null(input$posttest_file)) {
        df <- readxl::read_excel(input$posttest_file$datapath)
        
        # Pastikan jumlah kolom yang ada adalah 3 (Nama Peserta, Departemen, Nilai Post-Test)
        if (ncol(df) != 3) {
          upload_message("Error: File Post-Test harus memiliki 3 kolom.")
          upload_color("#b22222")  # Error: red
          return(NULL)
        }
        
        # Rename kolom sesuai dengan yang diinginkan
        colnames(df) <- c("Nama Peserta", "Departemen", "Nilai Post-Test")
        
        if (!all(c("Nama Peserta", "Departemen", "Nilai Post-Test") %in% colnames(df))) {
          upload_message("Error: File Post-Test tidak sesuai format.")
          upload_color("#b22222")  # Error: red
        }
      }
      
    }, error = function(e) {
      # Catch any errors and show them
      upload_message(paste("Error: ", e$message))
      upload_color("#b22222")  # Error: red
    })
    
    # Jika semua file valid
    if (!is.null(input$feedback_file) && !is.null(input$pretest_file) && !is.null(input$posttest_file)) {
      upload_message("Data berhasil diunggah. Silakan ke tab Report untuk melihat hasil.<br>Jika output tidak tampil sepenuhnya, pastikan file yang diunggah tidak tertukar.")
      upload_color("#006400")  
    } else {
      upload_message("File belum diunggah.")
      upload_color("#0056a1")  
    }
  })
  
  # Output the upload status message
  output$upload_status <- renderText({
    HTML(paste('<div style="color:', upload_color(), '; margin-top: 0px; margin-bottom: 10px; font-size:12px">', 
               upload_message(), '</div>'))
  })
  
  # Combine Data
  combined_data <- reactive({
    feedback <- feedback_data()
    pre <- pretest_data()
    post <- posttest_data()
    
    if (!is.null(feedback) && !is.null(pre) && !is.null(post)) {
      # Sort data by "Nama Peserta" (ascending order)
      feedback <- feedback %>% arrange(`Nama Peserta`)
      pre <- pre %>% arrange(`Nama Peserta`)
      post <- post %>% arrange(`Nama Peserta`)
      
      # Remove the "Nama Peserta" column from pre and post data
      pre <- pre %>% select(-`Nama Peserta`)
      post <- post %>% select(-`Nama Peserta`)
      
      # Combine all data
      combined <- bind_cols(feedback, pre, post)
      return(combined)
    } else {
      return(NULL)
    }
  })
  
  # Output Texts
  output$training_title <- renderText({
    combined <- combined_data()
    if (!is.null(combined) && "Judul Pelatihan" %in% colnames(combined)) {
      HTML(paste0(
        "<span style='font-size: 40px;font-weight: bold;'>",  # Atur ukuran dan tebal font
        combined$`Judul Pelatihan`[1],
        "</span>"
      ))
    } else {
      HTML("<span style='font-size: 40px;font-weight: bold;'>Judul Pelatihan: Tidak tersedia</span>")
    }
  })
  
  output$training_date <- renderText({
    combined <- combined_data()
    if (!is.null(combined) && "Timestamp" %in% colnames(combined)) {
      date <- as.Date(combined$Timestamp[1])
      HTML(paste0(
        "<span style='font-size: 24px;'>",  # Atur ukuran, tebal, dan warna font
        format(date, "%d %B %Y"),
        "</span>"
      ))
    } else {
      HTML("<span style='font-size: 24px;'>Tanggal Pelaksanaan: Tidak tersedia</span>")
    }
  })
  
  
  output$training_location <- renderText({
    combined <- combined_data()
    if (!is.null(combined) && "Tempat Pelaksanaan" %in% colnames(combined)) {
      HTML(paste0(
        "<span style='font-size: 24px;'>",  # Atur ukuran, tebal, dan warna font
        combined$`Tempat Pelaksanaan`[1],
        "</span>"
      ))
    } else {
      HTML("<span style='font-size: 24px;'>Tempat Pelaksanaan: Tidak tersedia</span>")
    }
  })
  
  # Output: Grafik Training Score
  output$training_fig <- renderPlot({
    feedback <- feedback_data()
    
    if (!is.null(feedback)) {
      # Menghitung persentase "Ya" dan "Tidak" untuk setiap pertanyaan
      training <- feedback %>%
        select(`Kejelasan Tujuan`, `Ketercapaian Tujuan`, `Kesesuaian Materi`,
               `Penerapan Materi`, `Tambahan Pengetahuan`) %>%
        summarise(across(everything(), 
                         list(Ya = ~ mean(. == "Ya", na.rm = TRUE) * 100,
                              Tidak = ~ mean(. == "Tidak", na.rm = TRUE) * 100)))
      
      # Konversi ke format panjang untuk stacked bar chart
      plot_data <- training %>%
        pivot_longer(cols = everything(), names_to = c("Question", "Category"),
                     names_sep = "_", values_to = "Percentage") %>%
        mutate(Question = recode(Question,
                                 `Kejelasan Tujuan` = "Kejelasan Tujuan",
                                 `Ketercapaian Tujuan` = "Ketercapaian Tujuan",
                                 `Kesesuaian Materi` = "Kesesuaian Materi",
                                 `Penerapan Materi` = "Penerapan Materi",
                                 `Tambahan Pengetahuan` = "Tambahan Pengetahuan"),
               Category = factor(Category, levels = c("Tidak", "Ya")),
               Text = paste0(round(Percentage, 1), "%"))  # Menambahkan teks persentase
      
      # Membuat ggplot2 stacked bar chart
      plot <- ggplot(plot_data, aes(x = Question, y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", position = "stack", width = 0.6) +  # Ukuran batang diperkecil
        geom_text(aes(label = ifelse(Percentage > 0, Text, NA)),  # Hanya tampilkan jika > 0
                  position = position_stack(vjust = 0.5), 
                  color = "white", size = 4) +  # Ukuran teks batang disesuaikan
        scale_fill_manual(values = c("Ya" = "#0056A1", "Tidak" = "#B22222")) +
        labs(x = "Questions", y = "Percentage (%)", fill = "Category") +
        theme_minimal(base_size = 12) +  # Ukuran dasar elemen diperbesar
        theme(
          axis.title.x = element_text(size = 14, margin = margin(t = 15)),  # Tambah jarak ke atas
          axis.title.y = element_text(size = 14, margin = margin(r = 15)),
          axis.text.x = element_text(size = 12, margin = margin(t = 10)),  # Tambah jarak ke bawah
          axis.text.y = element_text(size = 12), margin = margin(r = 10),
          legend.title = element_text(size = 14),
          legend.text = element_text(size = 12),
          plot.margin = margin(15, 15, 15, 15),  # Tambah margin
          panel.grid.minor = element_blank()   # Hilangkan garis grid kecil
        ) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.1)))  # Jarak lebih rapi di sumbu-y
      
      return(plot)
    }
  })
  
  # Grafik Fasilitator Score
  output$fasil_fig <- renderPlot({
    combined <- combined_data()
    
    if (!is.null(combined)) {
      # Menghitung persentase "Ya" dan "Tidak" untuk setiap pertanyaan fasilitator
      fasil <- combined %>% 
        select(`Kemudahan Pemahaman`, `Fasilitator Menguasai`, `Kompetensi Fasilitator`, 
               `Keaktifan Fasilitator`) %>%
        summarise(across(everything(), 
                         list(Ya = ~ mean(. == "Ya", na.rm = TRUE) * 100,
                              Tidak = ~ mean(. == "Tidak", na.rm = TRUE) * 100)))
      
      # Konversi ke format panjang untuk stacked bar chart
      plot_data <- fasil %>%
        pivot_longer(cols = everything(), names_to = c("Question", "Category"),
                     names_sep = "_", values_to = "Percentage") %>%
        mutate(Question = recode(Question,
                                 `Kemudahan Pemahaman` = "Kemudahan Pemahaman",
                                 `Fasilitator Menguasai` = "Fasilitator Menguasai",
                                 `Kompetensi Fasilitator` = "Kompetensi Fasilitator",
                                 `Keaktifan Fasilitator` = "Keaktifan Fasilitator"),
               Category = factor(Category, levels = c("Tidak", "Ya")),
               Text = paste0(round(Percentage, 1), "%"))  # Menambahkan teks persentase
      
      # Membuat ggplot2 stacked bar chart
      plot <- ggplot(plot_data, aes(x = Question, y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", position = "stack", width = 0.6) +  # Ukuran batang diperkecil
        geom_text(aes(label = ifelse(Percentage > 0, Text, NA)),  # Hanya tampilkan jika > 0
                  position = position_stack(vjust = 0.5), 
                  color = "white", size = 4) +  # Ukuran teks batang disesuaikan
        scale_fill_manual(values = c("Ya" = "#0056A1", "Tidak" = "#B22222")) +
        labs(x = "Questions", y = "Percentage (%)", fill = "Category") +
        theme_minimal(base_size = 12) +  # Ukuran dasar elemen diperbesar
        theme(
          axis.title.x = element_text(size = 14, margin = margin(t = 15)),  # Tambah jarak ke atas
          axis.title.y = element_text(size = 14, margin = margin(r = 15)),
          axis.text.x = element_text(size = 12, margin = margin(t = 10)),  # Tambah jarak ke bawah
          axis.text.y = element_text(size = 12, margin = margin(r = 10)),
          legend.title = element_text(size = 14),
          legend.text = element_text(size = 12),
          plot.margin = margin(15, 15, 15, 15),  # Tambah margin
          panel.grid.minor = element_blank()   # Hilangkan garis grid kecil
        ) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.1)))  # Jarak lebih rapi di sumbu-y
      
      return(plot)
    }
  })
  
  
  # Difficulty Level Pie Chart
  output$pie_chart <- renderPlot({
    data <- combined_data()
    
    if (!is.null(data)) {
      # Membersihkan data dan menghitung frekuensi
      diff_level <- na.omit(data$`Tingkat Kesulitan`)
      diff_frequencies <- table(factor(diff_level, levels = c(5, 4, 3, 2, 1)))
      
      # Label kategori
      custom_labels_diff <- c("Sangat Sulit", "Sulit", "Biasa Saja", "Mudah", "Sangat Mudah")
      
      # Data frame untuk ggplot2
      df <- data.frame(
        Category = custom_labels_diff,
        Frequency = as.numeric(diff_frequencies),
        Percentage = as.numeric(diff_frequencies) / sum(diff_frequencies) * 100  # Perhitungan persentase
      )
      
      # Warna untuk setiap kategori
      colors <- c("#B22222", "#FF8C00", "#ffd700", "#228b22", "#0056A1")
      
      # Membuat pie chart menggunakan ggplot2
      plot <- ggplot(df, aes(x = "", y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", width = 1, color = "white") +  # Membuat batang berbentuk lingkaran
        coord_polar(theta = "y") +  # Mengubah ke pie chart
        scale_fill_manual(values = colors, limits = custom_labels_diff) +  # Menerapkan warna kustom
        geom_text(aes(label = ifelse(Percentage > 0, paste0(round(Percentage, 1), "%"), "")), 
                  position = position_stack(vjust = 0.5), color = "white", size = 5) +  # Teks lebih besar
        theme_void(base_size = 10) +  # Base size diperbesar
        theme(
          legend.title = element_blank(),  # Hilangkan judul legenda
          legend.text = element_text(size = 10),
          legend.position = "right",  # Legenda di kanan
          legend.key.size = unit(1, "cm"),  # Ukuran legenda lebih besar
          plot.margin = margin(10, 10, 10, 10)  # Tambahkan margin
        )
      
      return(plot)
    }
  })
  
  # Fasil Score Pie Chart
  output$score_pie_chart <- renderPlot({
    data <- combined_data()
    
    if (!is.null(data)) {
      # Membersihkan data dan menghitung frekuensi
      fasil_score <- na.omit(data$`Penilaian Fasilitator`)
      diff_frequencies <- table(factor(fasil_score, levels = c(5, 4, 3, 2, 1)))
      
      # Label kategori
      custom_labels_diff <- c("Sangat Tinggi", "Tinggi", "Sedang", "Rendah", "Sangat Rendah")
      
      # Data frame untuk ggplot2
      df <- data.frame(
        Category = custom_labels_diff,
        Frequency = as.numeric(diff_frequencies),
        Percentage = as.numeric(diff_frequencies) / sum(diff_frequencies) * 100  # Perhitungan persentase
      )
      
      # Warna untuk setiap kategori
      colors <- c("#0056A1", "#228b22", "#ffd700", "#FF8C00", "#B22222")
      
      # Membuat pie chart menggunakan ggplot2
      plot <- ggplot(df, aes(x = "", y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", width = 1, color = "white") +  # Membuat batang berbentuk lingkaran
        coord_polar(theta = "y") +  # Mengubah ke pie chart
        scale_fill_manual(values = colors, limits = custom_labels_diff) +  # Menerapkan warna kustom
        geom_text(aes(label = ifelse(Percentage > 0, paste0(round(Percentage, 1), "%"), "")), 
                  position = position_stack(vjust = 0.5), color = "white", size = 5) +  # Teks lebih besar
        theme_void(base_size = 10) +  # Base size diperbesar
        theme(
          legend.title = element_blank(),  # Hilangkan judul legenda
          legend.text = element_text(size = 10),
          legend.position = "right",  # Legenda di kanan
          legend.key.size = unit(1, "cm"),  # Ukuran legenda lebih besar
          plot.margin = margin(10, 10, 10, 10)  # Tambahkan margin
        )
      
      return(plot)
    }
  })
  
  # Total Score Bar Chart
  output$total_score_fig <- renderPlot({
    data <- combined_data()
    
    if (!is.null(data)) {
      # Menentukan kategori dan urutannya
      data <- data %>%
        mutate(
          Total_Penilaian_Kategori = case_when(
            `Total Penilaian` == 5 ~ "Sangat Tinggi",
            `Total Penilaian` == 4 ~ "Tinggi",
            `Total Penilaian` == 3 ~ "Sedang",
            `Total Penilaian` == 2 ~ "Rendah",
            `Total Penilaian` == 1 ~ "Sangat Rendah"
          )
        )
      
      # Menghitung frekuensi setiap kategori
      freq_data <- data %>%
        group_by(Total_Penilaian_Kategori) %>%
        summarise(Frequency = n(), .groups = 'drop') %>%
        # Menambahkan kategori yang tidak muncul dengan level yang sudah ditentukan
        mutate(Total_Penilaian_Kategori = factor(
          Total_Penilaian_Kategori,
          levels = c("Sangat Tinggi", "Tinggi", "Sedang", "Rendah", "Sangat Rendah")
        )) %>%
        complete(Total_Penilaian_Kategori, fill = list(Frequency = 0)) %>%
        arrange(Total_Penilaian_Kategori)
      
      # Warna kustom
      colors <- c(
        "Sangat Tinggi" = "#003366",  # Biru tua
        "Tinggi" = "#228B22",         # Hijau gelap
        "Sedang" = "#FFD700",         # Kuning biasa
        "Rendah" = "#B22222",         # Merah gelap
        "Sangat Rendah" = "#8B0000"  # Merah marun
      )
      
      # Membuat bar chart horizontal
      plot <- ggplot(freq_data, aes(x = Frequency, y = Total_Penilaian_Kategori, fill = Total_Penilaian_Kategori)) +
        geom_bar(stat = "identity", color = "white", width = 0.7) +  # Bar chart horizontal
        scale_fill_manual(values = colors) +  # Terapkan warna kustom
        scale_y_discrete(limits = rev(levels(freq_data$Total_Penilaian_Kategori))) +  # Membalik urutan sumbu Y
        geom_text(aes(label = Frequency), hjust = -0.2, color = "white", size = 5) +  # Label frekuensi
        theme_minimal(base_size = 12) +  # Tema minimal
        theme(
          legend.position = "none",  # Hilangkan legenda
          axis.title.y = element_text(size = 12, margin = margin(r = 10)),  # Judul sumbu Y
          axis.title.x = element_text(size = 12, margin = margin(t = 10)),  # Judul sumbu X
          axis.text.y = element_text(size = 10),  # Ukuran teks sumbu Y
          axis.text.x = element_text(size = 10),  # Ukuran teks sumbu X
          plot.margin = margin(20, 20, 20, 20)  # Margin grafik
        ) +
        labs(
          x = "Frekuensi",  # Judul sumbu X
          y = "Kategori Penilaian"  # Judul sumbu Y
        )
      
      return(plot)
    }
  })
  
  
  # Preprocessing function
  preprocess_text <- function(text) {
    text <- str_replace_all(text, "(^|\\s)(\\d+\\.|[-•*#])\\s*", "")  
    text <- str_replace_all(text, "\\s{2,}", ". ")                      
    text <- str_replace_all(text, "\\s*\\.\\s*", ". ")                 
    text <- str_replace_all(text, "\\s*\\r?\\n\\s*", ", ")             
    text <- str_replace_all(text, "(?<!\\.)\\s*,\\s*(?!\\.)", ", ")    
    text <- str_trim(text)                                             
    text <- str_replace_all(text, "(?<!\\.)$", ".")
    
    return(text)
  }
  
  
  # Function to extract the 3 longest sentences
  process_text <- function(df, text_column) {
    df_clean <- df %>% filter(!is.na(.data[[text_column]]))  # Remove NAs
    
    # Apply preprocessing to clean the text first
    df_clean <- df_clean %>%
      mutate(cleaned_text = sapply(.data[[text_column]], preprocess_text),  # Apply preprocess_text here
             length = nchar(cleaned_text)) %>%  # Get the length of the cleaned text
      arrange(desc(length)) %>%  # Sort by length (longest first)
      slice(2:4)  # Take the top 3 longest entries
    
    return(df_clean$cleaned_text)  # Return the processed and top 3 sentences
  }
  
  # Summary
  output$material_summary <- renderUI({
    req(combined_data())
    top_material_sentences <- process_text(combined_data(), "Point Utama Materi")
    
    # Teks yang dapat diedit
    textAreaInput(
      inputId = "editable_material_summary",
      value = paste("•", top_material_sentences, collapse = "\n"),
      label = NULL,
      width = "100%",
      height = "200px"
    )
  })
  
  output$suggestion_summary <- renderUI({
    req(combined_data())
    top_suggestion_sentences <- process_text(combined_data(), "Kritik dan Saran")
    
    # Teks yang dapat diedit
    textAreaInput(
      inputId = "editable_suggestion_summary",
      value = paste("•", top_suggestion_sentences, collapse = "\n"),
      label = NULL,
      width = "100%",
      height = "200px"
    )
  })
  
  observeEvent(input$editable_material_summary, {
    print(input$editable_material_summary)  # Menangkap teks yang telah diedit
  })
  
  observeEvent(input$editable_suggestion_summary, {
    print(input$editable_suggestion_summary)  # Menangkap teks yang telah diedit
  })
  
  
  # Output tabel
  output$pre_post_table <- renderTable({
    combined <- combined_data()
    
    table_data <- combined %>%
      mutate(No = row_number(),
             `Nilai Pre-Test` = as.integer(round(`Nilai Pre-Test`, 0)),  # Konversi ke bilangan bulat
             `Nilai Post-Test` = as.integer(round(`Nilai Post-Test`, 0)), # Konversi ke bilangan bulat
             Gain = as.integer(round(`Nilai Post-Test` - `Nilai Pre-Test`, 0))) %>% # Konversi ke bilangan bulat
      select(No, `Nama Peserta`, `Nilai Pre-Test`, `Nilai Post-Test`, Gain) %>%
      arrange(`Nama Peserta`) # Sort by participant's name
    
    table_data
  }, sanitize.text.function = function(x) x)  # Hindari perubahan format angka
  
  
  
  # Apply custom CSS for styling
  output$table_styles <- renderUI({
    tags$style(HTML("
    .table {
      width: 100%;
      max-width: 1200px; /* Membatasi lebar tabel */
      margin: 0 auto; /* Memusatkan tabel */
      background-color: white;
      font-size: 16px;
      border-collapse: collapse;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .table th {
      font-weight: bold;
      text-align: center;
      background-color: #d3d3d3; /* Abu-abu untuk header */
      color: black; /* Warna teks di header */
      padding: 12px;
    }
    .table td {
      text-align: center;
      padding: 12px;
      background-color: #f9f9f9; /* Warna dasar tabel */
    }
    .table tr:nth-child(even) td {
      background-color: #f1f1f1; /* Striping untuk baris */
    }
    .table td:nth-child(2) {
      text-align: left; /* Rata kiri untuk kolom 'Nama Peserta' */
      padding-left: 16px; /* Tambahkan jarak kiri */
    }
    .table th, .table td {
      border: 1px solid #ddd;
    }
    @media (max-width: 2000px) {
      .table {
        font-size: 14px; /* Ukuran lebih kecil untuk layar kecil */
      }
    }
  "))
  })
  
  
  # Output boxplot
  output$boxplot <- renderPlot({
    req(combined_data())
    
    combined <- combined_data()
    
    # Reshaping the data for ggplot
    data_nilai <- combined %>%
      select(`Nama Peserta`, `Nilai Pre-Test`, `Nilai Post-Test`) %>%
      pivot_longer(cols = starts_with("Nilai"), names_to = "Kategori", values_to = "Nilai") %>%
      mutate(
        Kategori = recode(Kategori, "Nilai Pre-Test" = "Pre-Test", "Nilai Post-Test" = "Post-Test"),
        Kategori = factor(Kategori, levels = c("Pre-Test", "Post-Test")) # Ensure Pre-Test is on the left
      )
    
    # Custom colors
    colors <- c("Pre-Test" = "#0056A1", "Post-Test" = "#0056A1")
    
    # Creating the boxplot with ggplot2
    plot <- ggplot(data_nilai, aes(x = Kategori, y = Nilai, fill = Kategori)) +
      geom_boxplot(outlier.color = "red", outlier.shape = 21, outlier.size = 2) +  # Customize outlier appearance
      scale_fill_manual(values = colors) +  # Set custom colors
      theme_minimal(base_size = 15) +  # Minimal theme with larger font size
      theme(
        legend.position = "none",  # Remove legend
        axis.title.x = element_text(size = 14, margin = margin(t = 10)),  # X-axis title style
        axis.title.y = element_text(size = 14, margin = margin(r = 10)),  # Y-axis title style
        axis.text = element_text(size = 12),  # Axis text size
        plot.margin = margin(20, 20, 20, 20)  # Adjust margins
      ) +
      labs(
        x = "Kategori",  # X-axis label
        y = "Nilai"      # Y-axis label
      )
    
    return(plot)
  })
  
  
  output$stat_result <- renderUI({
    req(combined_data())
    
    t_test_result <- t.test(combined_data()$`Nilai Post-Test`, combined_data()$`Nilai Pre-Test`, paired = TRUE, alternative = "greater")
    p_value_one_tailed <- t_test_result$p.value
    
    result_text <- ""
    if (p_value_one_tailed < 0.05) {
      result_text <- paste("Pengujian secara statistik menunjukkan hasil yang signifikan bahwa rata-rata nilai post-test lebih tinggi dibandingkan dengan nilai pre-test karena memiliki tingkat kesalahan sebesar", 
                           round(p_value_one_tailed * 100, 2), "% yang lebih kecil dari 5%.")
    } else {
      result_text <- paste("Pengujian secara statistik tidak menunjukkan hasil yang signifikan bahwa rata-rata nilai post-test lebih tinggi dibandingkan dengan nilai pre-test karena memiliki tingkat kesalahan sebesar", 
                           round(p_value_one_tailed * 100, 2), "% yang lebih besar dari 5%.")
    }
    
    tagList(
      div(
        class = "box-custom",
        style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
        div(
          h4(style = "font-size: 22px; color: #333; font-weight: bold; text-align: left;", "Apakah ada peningkatan nilai setelah pelatihan?"),
          h3(style = "font-size: 16px; color: #333; text-align: left;", result_text) # Display the result text inside a div
        )
      )
    )
  })
  
  
  output$corr_result <- renderUI({
    req(combined_data())
    
    # Perform Spearman correlation test
    correlation_result <- cor.test(combined_data()$`Tingkat Kesulitan`, combined_data()$`Nilai Post-Test`, method = "spearman")
    
    result_text <- ""
    
    # Check if correlation is significant
    if (correlation_result$p.value < 0.05) {
      # Handle negative or positive correlation
      if (correlation_result$estimate < 0) {
        result_text <- paste("Secara statistik, terdapat pengaruh antara tingkat kesulitan materi dengan nilai post-test yang didapatkan. Korelasi negatif (", 
                             round(correlation_result$estimate, 2), 
                             ") menunjukkan bahwa peserta yang menganggap materi lebih sulit cenderung memiliki nilai post-test yang lebih rendah.")
      } else {
        result_text <- paste("Secara statistik, terdapat pengaruh antara tingkat kesulitan materi dengan nilai post-test yang didapatkan. Korelasi positif (", 
                             round(correlation_result$estimate, 2), 
                             ") menunjukkan bahwa peserta yang menganggap materi lebih sulit cenderung memiliki nilai post-test yang lebih tinggi.")
      }
    } else {
      result_text <- "Tidak terdapat pengaruh yang signifikan antara tingkat kesulitan materi dengan nilai post-test yang didapatkan."
    }
    
    # Create UI layout with styled box
    tagList(
      div(
        class = "box-custom",
        style = "margin: 20px auto; padding: 20px; width: 100%; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); border-radius: 15px;",
        div(
          h4(style = "font-size: 22px; color: #333; font-weight: bold; text-align: left;", "Apakah ada pengaruh antara nilai post-test dengan pilihan tingkat kesulitan materi oleh peserta?"),
          h3(style = "font-size: 16px; color: #333; text-align: left;", result_text) # Display the result text inside a div
        )
      )
    )
  })
  
  output$download_button <- downloadHandler(
    filename = function() {
      combined <- combined_data()  # Call the reactive expression to get the data
      training_title <- combined$`Judul Pelatihan`[1]  # Access the title from the reactive data
      paste("TrainingEvaluation_", training_title, "_", Sys.Date(), ".docx", sep = "") 
    },
    content = function(file) {
      combined <- combined_data()
      print(names(combined))
      if (is.null(combined)) {
        showNotification("Error: Data gabungan tidak tersedia", type = "error")
        return(NULL)
      }
      
      
      
      # Grafik Training
      tf_data <- combined %>%
        select(`Kejelasan Tujuan`, `Ketercapaian Tujuan`, `Kesesuaian Materi`,
               `Penerapan Materi`, `Tambahan Pengetahuan`) %>%
        summarise(across(everything(), 
                         list(Ya = ~ mean(. == "Ya", na.rm = TRUE) * 100,
                              Tidak = ~ mean(. == "Tidak", na.rm = TRUE) * 100))) %>%
        pivot_longer(cols = everything(), names_to = c("Question", "Category"),
                     names_sep = "_", values_to = "Percentage") %>%
        mutate(Question = recode(Question,
                                 `Kejelasan Tujuan` = "Kejelasan Tujuan",
                                 `Ketercapaian Tujuan` = "Ketercapaian Tujuan",
                                 `Kesesuaian Materi` = "Kesesuaian Materi",
                                 `Penerapan Materi` = "Penerapan Materi",
                                 `Tambahan Pengetahuan` = "Tambahan Pengetahuan"),
               Category = factor(Category, levels = c("Tidak", "Ya")),
               Text = paste0(round(Percentage, 1), "%")) %>%
        # Menambahkan kategori "Tidak" jika tidak ada dalam data
        { 
          if (!"Tidak" %in% levels(. $Category)) {
            empty_row <- data.frame(
              Question = unique(. $Question),
              Category = "Tidak",
              Percentage = 0,
              Text = "0%"
            )
            bind_rows(., empty_row)
          } else {
            .
          }
        }
      
      plot_tf <- ggplot(tf_data, aes(x = Question, y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", position = "stack", width = 0.6) +  # Ukuran batang diperkecil
        geom_text(aes(label = ifelse(Percentage > 0, Text, NA)),  # Hanya tampilkan jika > 0
                  position = position_stack(vjust = 0.5), 
                  color = "white", size = 4) +  # Ukuran teks batang disesuaikan
        scale_fill_manual(values = c("Ya" = "#0056A1", "Tidak" = "#B22222"), breaks = c("Ya", "Tidak")) +
        labs(x = "Questions", y = "Percentage (%)", fill = "Category") +
        theme_minimal(base_size = 12) +  # Ukuran dasar elemen diperbesar
        theme(
          axis.title.x = element_text(size = 14, margin = margin(t = 15)),  # Tambah jarak ke atas
          axis.title.y = element_text(size = 14, margin = margin(r = 15)),
          axis.text.x = element_text(size = 12, margin = margin(t = 10)),  # Tambah jarak ke bawah
          axis.text.y = element_text(size = 12), margin = margin(r = 10),
          legend.title = element_text(size = 14),
          legend.text = element_text(size = 12),
          plot.margin = margin(15, 15, 15, 15),  # Tambah margin
          panel.grid.minor = element_blank()   # Hilangkan garis grid kecil
        ) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.1)))
      
      
      fasil <- combined %>% 
        select(`Kemudahan Pemahaman`, `Fasilitator Menguasai`, `Kompetensi Fasilitator`, 
               `Keaktifan Fasilitator`) %>%
        summarise(across(everything(), 
                         list(Ya = ~ mean(. == "Ya", na.rm = TRUE) * 100,
                              Tidak = ~ mean(. == "Tidak", na.rm = TRUE) * 100)))
      
      # Konversi ke format panjang untuk stacked bar chart
      ff_data <- fasil %>%
        pivot_longer(cols = everything(), names_to = c("Question", "Category"),
                     names_sep = "_", values_to = "Percentage") %>%
        mutate(Question = recode(Question,
                                 `Kemudahan Pemahaman` = "Kemudahan Pemahaman",
                                 `Fasilitator Menguasai` = "Fasilitator Menguasai",
                                 `Kompetensi Fasilitator` = "Kompetensi Fasilitator",
                                 `Keaktifan Fasilitator` = "Keaktifan Fasilitator"),
               Category = factor(Category, levels = c("Tidak", "Ya")),
               Text = paste0(round(Percentage, 1), "%"))  # Menambahkan teks persentase
      
      # Membuat ggplot2 stacked bar chart
      plot_ff <- ggplot(ff_data, aes(x = Question, y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", position = "stack", width = 0.6) +
        geom_text(aes(label = ifelse(Percentage > 0, Text, NA)), 
                  position = position_stack(vjust = 0.5), color = "white", size = 4) +
        scale_fill_manual(values = c("Ya" = "#0056A1", "Tidak" = "#B22222")) +
        labs(x = "Questions", y = "Percentage (%)", fill = "Category") +
        theme_minimal(base_size = 12) +
        theme(
          axis.title.x = element_text(size = 14, margin = margin(t = 15)),
          axis.title.y = element_text(size = 14, margin = margin(r = 15)),
          axis.text.x = element_text(size = 12, margin = margin(t = 10)),
          axis.text.y = element_text(size = 12, margin = margin(r = 10)),
          legend.title = element_text(size = 14),
          legend.text = element_text(size = 12),
          plot.margin = margin(15, 15, 15, 15),
          panel.grid.minor = element_blank()
        ) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.1))) 
      
      # Diff Level Pie Chart
      diff_level <- na.omit(combined$`Tingkat Kesulitan`)
      diff_frequencies <- table(factor(diff_level, levels = c(5, 4, 3, 2, 1)))
      
      # Label kategori
      custom_labels_diff <- c("Sangat Sulit", "Sulit", "Biasa Saja", "Mudah", "Sangat Mudah")
      
      # Data frame untuk ggplot2
      df_diff <- data.frame(
        Category = custom_labels_diff,
        Frequency = as.numeric(diff_frequencies),
        Percentage = as.numeric(diff_frequencies) / sum(diff_frequencies) * 100  # Perhitungan persentase
      )
      
      # Warna untuk setiap kategori
      colors <- c("#B22222", "#FF8C00", "#ffd700", "#228b22", "#0056A1")
      
      # Membuat pie chart menggunakan ggplot2
      plot_diff <- ggplot(df_diff, aes(x = "", y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", width = 1, color = "white") + 
        coord_polar(theta = "y") + 
        scale_fill_manual(values = colors, limits = custom_labels_diff) + 
        geom_text(aes(label = ifelse(Percentage > 0, paste0(round(Percentage, 1), "%"), "")), 
                  position = position_stack(vjust = 0.5), color = "white", size = 5) +
        theme_void(base_size = 10) + 
        theme(
          legend.title = element_blank(), 
          legend.text = element_text(size = 10),
          legend.position = "right", 
          legend.key.size = unit(1, "cm"),
          plot.margin = margin(10, 10, 10, 10)
        )
      
      # Fasil Score Pie
      fasil_score <- na.omit(combined$`Penilaian Fasilitator`)
      fasil_frequencies <- table(factor(fasil_score, levels = c(5, 4, 3, 2, 1)))
      
      custom_labels_fasil <- c("Sangat Tinggi", "Tinggi", "Sedang", "Rendah", "Sangat Rendah")
      
      df_fasil <- data.frame(
        Category = custom_labels_fasil,
        Frequency = as.numeric(fasil_frequencies),
        Percentage = as.numeric(fasil_frequencies) / sum(fasil_frequencies) * 100
      )
      
      colors_fasil <- c("#0056A1", "#228b22", "#ffd700", "#FF8C00", "#B22222")
      
      plot_fasil <- ggplot(df_fasil, aes(x = "", y = Percentage, fill = Category)) +
        geom_bar(stat = "identity", width = 1, color = "white") + 
        coord_polar(theta = "y") + 
        scale_fill_manual(values = colors_fasil, limits = custom_labels_fasil) + 
        geom_text(aes(label = ifelse(Percentage > 0, paste0(round(Percentage, 1), "%"), "")), 
                  position = position_stack(vjust = 0.5), color = "white", size = 5) +
        theme_void(base_size = 10) + 
        theme(
          legend.title = element_blank(), 
          legend.text = element_text(size = 10),
          legend.position = "right", 
          legend.key.size = unit(1, "cm"),
          plot.margin = margin(10, 10, 10, 10)
        )
      
      # Membuat Total Score Bar Chart
      total_score_data <- combined %>%
        mutate(
          Total_Penilaian_Kategori = case_when(
            `Total Penilaian` == 5 ~ "Sangat Tinggi",
            `Total Penilaian` == 4 ~ "Tinggi",
            `Total Penilaian` == 3 ~ "Sedang",
            `Total Penilaian` == 2 ~ "Rendah",
            `Total Penilaian` == 1 ~ "Sangat Rendah"
          )
        ) %>%
        group_by(Total_Penilaian_Kategori) %>%
        summarise(Frequency = n(), .groups = 'drop') %>%
        mutate(Total_Penilaian_Kategori = factor(
          Total_Penilaian_Kategori,
          levels = c("Sangat Tinggi", "Tinggi", "Sedang", "Rendah", "Sangat Rendah")
        )) %>%
        complete(Total_Penilaian_Kategori, fill = list(Frequency = 0)) %>%
        arrange(Total_Penilaian_Kategori)
      
      colors_score <- c(
        "Sangat Tinggi" = "#003366",
        "Tinggi" = "#228B22",
        "Sedang" = "#FFD700",
        "Rendah" = "#B22222",
        "Sangat Rendah" = "#8B0000"
      )
      
      plot_score <- ggplot(total_score_data, aes(x = Frequency, y = Total_Penilaian_Kategori, fill = Total_Penilaian_Kategori)) +
        geom_bar(stat = "identity", color = "white", width = 0.7) + 
        scale_fill_manual(values = colors_score) +  
        scale_y_discrete(limits = rev(levels(total_score_data$Total_Penilaian_Kategori))) +  
        geom_text(aes(label = Frequency), hjust = -0.2, color = "white", size = 5) +  
        theme_minimal(base_size = 12) +  
        theme(
          legend.position = "none",  
          axis.title.y = element_text(size = 12, margin = margin(r = 10)),
          axis.title.x = element_text(size = 12, margin = margin(t = 10)),
          axis.text.y = element_text(size = 10),
          axis.text.x = element_text(size = 10),
          plot.margin = margin(20, 20, 20, 20)
        ) +
        labs(x = "Frekuensi", y = "Kategori Penilaian")
      
      
      temp_file_training <- tempfile(fileext = ".png")
      temp_file_fasil <- tempfile(fileext = ".png")
      temp_file_diff <- tempfile(fileext = ".png")
      temp_file_fasilscore <- tempfile(fileext = ".png")
      temp_file_totalscore <- tempfile(fileext = ".png")
      
      ggsave(temp_file_training, plot = plot_tf, width = 12, height = 4)
      ggsave(temp_file_fasil, plot = plot_ff, width = 12, height = 4)
      ggsave(temp_file_diff, plot = plot_diff, width = 6, height = 6)
      ggsave(temp_file_fasilscore, plot = plot_fasil, width = 6, height = 6)
      ggsave(temp_file_totalscore, plot = plot_score, width = 12, height = 4)
      
      # Materi dan Saran
      material_summary <- input$editable_material_summary
      suggestion_summary <- input$editable_suggestion_summary 
      
      # Tabel Data
      table_data <- combined %>%
        mutate(No = row_number(),
               `Nilai Pre-Test` = as.integer(round(`Nilai Pre-Test`, 0)),  
               `Nilai Post-Test` = as.integer(round(`Nilai Post-Test`, 0)),
               Gain = as.integer(round(`Nilai Post-Test` - `Nilai Pre-Test`, 0))) %>%
        select(No, `Nama Peserta`, `Nilai Pre-Test`, `Nilai Post-Test`, Gain) %>%
        arrange(`Nama Peserta`)
      
      pre_post_data <- flextable(table_data) %>%
        theme_vanilla() %>%
        autofit() %>%
        set_table_properties(layout = "autofit")
      

      # Box Plot Tabel
      data_nilai <- combined %>%
        select(`Nama Peserta`, `Nilai Pre-Test`, `Nilai Post-Test`) %>%
        pivot_longer(cols = starts_with("Nilai"), names_to = "Kategori", values_to = "Nilai") %>%
        mutate(
          Kategori = recode(Kategori, "Nilai Pre-Test" = "Pre-Test", "Nilai Post-Test" = "Post-Test"),
          Kategori = factor(Kategori, levels = c("Pre-Test", "Post-Test")) # Ensure Pre-Test is on the left
        )
      
      # Custom colors
      colors_boxplot <- c("Pre-Test" = "#0056A1", "Post-Test" = "#0056A1")
      
      plot_box <- ggplot(data_nilai, aes(x = Kategori, y = Nilai, fill = Kategori)) +
        geom_boxplot(outlier.color = "#0056a1", outlier.shape = 21, outlier.size = 2) + 
        scale_fill_manual(values = colors_boxplot) +
        theme_minimal(base_size = 15) +
        theme(
          legend.position = "none",
          axis.title.x = element_text(size = 14, margin = margin(t = 10)),
          axis.title.y = element_text(size = 14, margin = margin(r = 10)),
          axis.text = element_text(size = 12),
          plot.margin = margin(20, 20, 20, 20)
        ) +
        labs(
          x = "Kategori",  
          y = "Nilai"
        )
      
      temp_file_boxplot <- tempfile(fileext = ".png")
      ggsave(temp_file_boxplot, plot = plot_box, width = 4, height = 4)
      
      # Stat Result
      # Perform T-test (paired)
      t_test_result <- t.test(combined$`Nilai Post-Test`, combined$`Nilai Pre-Test`, paired = TRUE, alternative = "greater")
      p_value_one_tailed <- t_test_result$p.value
      
      result_text_ttest <- ""
      if (p_value_one_tailed < 0.05) {
        result_text_ttest <- paste("Pengujian secara statistik menunjukkan hasil yang signifikan bahwa rata-rata nilai post-test lebih tinggi dibandingkan dengan nilai pre-test karena memiliki tingkat kesalahan sebesar", 
                                   round(p_value_one_tailed * 100, 2), 
                                   "% yang lebih kecil dari 5%.")
      } else {
        result_text_ttest <- paste("Pengujian secara statistik tidak menunjukkan hasil yang signifikan bahwa rata-rata nilai post-test lebih tinggi dibandingkan dengan nilai pre-test karena memiliki tingkat kesalahan sebesar", 
                                   round(p_value_one_tailed * 100, 2), 
                                   "% yang lebih besar dari 5%.")
      }
      
      # Correlaation Result
      correlation_result <- cor.test(combined$`Tingkat Kesulitan`, combined$`Nilai Post-Test`, method = "spearman")
      
      result_text <- ""
      
      # Check if correlation is significant
      if (correlation_result$p.value < 0.05) {
        # Handle negative or positive correlation
        if (correlation_result$estimate < 0) {
          result_text <- paste("Secara statistik, terdapat pengaruh antara tingkat kesulitan materi dengan nilai post-test yang didapatkan. Korelasi negatif (", 
                               round(correlation_result$estimate, 2), 
                               ") menunjukkan bahwa peserta yang menganggap materi lebih sulit cenderung memiliki nilai post-test yang lebih rendah.")
        } else {
          result_text <- paste("Secara statistik, terdapat pengaruh antara tingkat kesulitan materi dengan nilai post-test yang didapatkan. Korelasi positif (", 
                               round(correlation_result$estimate, 2), 
                               ") menunjukkan bahwa peserta yang menganggap materi lebih sulit cenderung memiliki nilai post-test yang lebih tinggi.")
        }
      } else {
        result_text <- "Tidak terdapat pengaruh yang signifikan antara tingkat kesulitan materi dengan nilai post-test yang didapatkan."
      }
      
      
      # Membuat dokumen Word
      doc <- read_docx() %>%
        body_add_par("Training Evaluation", style = "centered") %>%
        body_add_par(value = combined$`Judul Pelatihan`[1], style = "centered") %>%
        body_add_par(value = format(as.Date(combined$Timestamp[1]), "%d %B %Y"), style = "centered") %>%
        body_add_par(value = combined$`Tempat Pelaksanaan`[1], style = "centered") %>%
        body_add_par("") %>%
        body_add_par("Training Feedback Visualization", style = "heading 1") %>%
        body_add_par("") %>%
        body_add_par("Training Score Bar Chart", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_training, width = 6, height = 2) %>%
        body_add_par("") %>%
        body_add_par("Fasilitator Score Bar Chart", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_fasil, width = 6, height = 2) %>%
        body_add_par("") %>%
        body_add_par("Difficulty Level Pie Chart", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_diff, width = 3, height = 3) %>%
        body_add_par("") %>%
        body_add_par("Fasilitator Assesment Pie Chart", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_fasilscore, width = 3, height = 3) %>%
        body_add_par("") %>%
        body_add_par("Total Training Score Bar Chart", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_totalscore, width = 6, height = 2) %>%
        body_add_par("") %>%
        body_add_par("Ringkasan Materi", style = "heading 2") %>%
        body_add_par(material_summary, style = "Normal") %>%
        body_add_par("Ringkasan Kritik dan Saran", style = "heading 2") %>%
        body_add_par(suggestion_summary, style = "Normal") %>%
        body_add_par("Laporan Pre dan Post Test", style = "heading 1") %>%
        body_add_par("") %>%
        body_add_par("Berikut adalah tabel hasil pre-test dan post-test:", style = "Normal") %>%
        body_add_flextable(pre_post_data) %>%
        body_add_par("Box Plot Pre-Test dan Post-Test", style = "heading 2") %>%
        body_add_par("") %>%
        body_add_img(src = temp_file_boxplot, width = 4, height = 4) %>%
        body_add_par("") %>%
        body_add_par("Apakah ada peningkatan nilai setelah pelatihan?", style = "heading 2") %>%
        body_add_par(result_text_ttest, style = "Normal") %>%
        body_add_par("Apakah ada pengaruh antara nilai post-test dengan pilihan tingkat kesulitan materi oleh peserta?", style = "heading 2") %>%
        body_add_par(result_text, style = "Normal")
      
      print(doc, target = file)
    }
  )
 
}

# Run the app
shinyApp(ui, server)

