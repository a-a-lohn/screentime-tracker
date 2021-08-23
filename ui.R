library(shiny)

fluidPage(
  headerPanel('ScreenTime Tracker'),
  sidebarPanel(
    fileInput('files', 'StayFree Export - Total Usage file', multiple = TRUE,
              accept = c('.xlsx', 'xls')),
    uiOutput('numApps'),
    uiOutput('dateRange')
    ),
  mainPanel(
    # plotOutput('plot1'),
    plotOutput('plot4'),
    plotOutput('plot3'),
    plotOutput('plot2'),
    plotOutput('pca'),
    plotOutput('contrib')
  )
)
