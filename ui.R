library(shiny)

fluidPage(
  headerPanel('Screentime Tracker'),
  sidebarPanel(
    fileInput('file', 'StayFree Export - Total Usage file',
              accept = c('.xlsx', 'xls')),
    uiOutput('numApps'),
    uiOutput('startDate'),
    uiOutput('endDate')
    ),
  # WHAT IF USER HAS <20 APPS?
  #sliderInput('numApps', 'Number of apps', 10, min = 1, max = 20),
  #uiOutput('dateRange')
  # dateRangeInput('dateRange', 'Date range', start = Sys.Date()-7, end = Sys.Date(),
  #                min = )
  mainPanel(
    # plotOutput('plot1'),
    plotOutput('plot4'),
    plotOutput('plot2'),
    plotOutput('plot3'),
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
