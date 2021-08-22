library(shiny)

fluidPage(
  headerPanel('ScreenTime Tracker'),
  sidebarPanel(
    fileInput('files', 'StayFree Export - Total Usage file', multiple = TRUE,
              accept = c('.xlsx', 'xls')),
    uiOutput('numApps'),
    #uiOutput('startDate'),
    #uiOutput('endDate')
    uiOutput('dateRange')
    ),
  #sliderInput('numApps', 'Number of apps', 10, min = 1, max = 20),
  #uiOutput('dateRange')
  # dateRangeInput('dateRange', 'Date range', start = Sys.Date()-7, end = Sys.Date(),
  #                min = )
  mainPanel(
    # plotOutput('plot1'),
    plotOutput('plot4'),
    plotOutput('plot3'),
    plotOutput('plot2'),
    plotOutput('pca'),
    plotOutput('contrib')
  )
)

#TODO:
# Add slider to filter dates and restrict dates according to files - DONE
# Allow multiple files to be uploaded
# Add ability to select specific apps instead of most popular ones
# Separate code into app.R and codebase file
# Upload to server
# Add note that it takes time to load file
# Change titles/names of elements in plots
# Add sample data to repo
# Improve aesthetics
