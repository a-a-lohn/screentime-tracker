library(shiny)

fluidPage(
  headerPanel('Screentime Tracker'),
  sidebarPanel(
    sliderInput('numApps', 'Number of apps', 10, min = 1, max = 20),
    fileInput('file', 'StayFree Export - Total Usage file',
              accept = c('.xlsx', 'xls'))
  ),
  mainPanel(
    # plotOutput('plot1'),
    plotOutput('plot4'),
    plotOutput('plot2'),
    plotOutput('plot3'),
    plotOutput('pca'),
    plotOutput('contrib')
  )
)