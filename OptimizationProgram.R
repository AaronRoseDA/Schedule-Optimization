rm(list = ls())
library(dplyr)
library(readxl)
library(lubridate)
library(ggplot2)
library(tidyr)
library(data.table)

dataMaster <- read_excel('C:/Users/arose/OneDrive - Ortho Molecular Products/Desktop/Schedule_Optimization/CurrentFormulations.xlsx',sheet = 'Export') 

create_empyty_room_Schedule <- function(){
  roomSchedule <- data.frame(
    time = format(seq(from = as.POSIXct("04:00:00", format="%H:%M:%S"), by = "1 min", length.out = 20*60), "%H:%M"),
    North = "Idle",
    South = "Idle",
    Old = "Idle",
    New = "Idle"
  )
  return(roomSchedule)
}

prepData <- function(){
  data <- dataMaster %>%
    mutate(
      mixRoom =
        case_when(
          grepl('1000', BatchSize, fixed = TRUE) ~ ifelse(round(runif(n(), 1, 2)) == 1, "North", "South"),
          grepl('100', BatchSize, fixed = TRUE) ~ ifelse(round(runif(n(), 1, 2)) == 1, "North", "South"),
          grepl('5000', BatchSize, fixed = TRUE) ~ ifelse(round(runif(n(), 3, 4)) == 3, "Old", "New"),
          grepl('500', BatchSize, fixed = TRUE) ~ ifelse(round(runif(n(), 1, 2)) == 1, "North", "South"),
          grepl('200', BatchSize, fixed = TRUE) ~ ifelse(round(runif(n(), 1, 2)) == 1, "North", "South"),
          TRUE ~ NA_character_
        ),
      millRoom =
        case_when(
          (mixRoom == "North" | mixRoom == "South") ~ mixRoom,
          (mixRoom == "Old" | mixRoom == "New") ~ ifelse(round(runif(n(), 1, 2)) == 1, "North", "South"),
          TRUE ~ NA_character_
        ),
      runOrder = sample(1:nrow(dataMaster), nrow(dataMaster), replace = F)
    ) %>%
    arrange(runOrder)
  return(data)
}

prepData <- function(){
  # Create a vectorized version of the mixRoom decision logic
  get_mix_room <- function(batch_size) {
    n <- length(batch_size)
    ifelse(grepl('1000', batch_size, fixed = TRUE), ifelse(round(runif(n, 1, 2)) == 1, "North", "South"),
           ifelse(grepl('100', batch_size, fixed = TRUE), ifelse(round(runif(n, 1, 2)) == 1, "North", "South"),
                  ifelse(grepl('5000', batch_size, fixed = TRUE), ifelse(round(runif(n, 3, 4)) == 3, "Old", "New"),
                         ifelse(grepl('500', batch_size, fixed = TRUE), ifelse(round(runif(n, 1, 2)) == 1, "North", "South"),
                                ifelse(grepl('200', batch_size, fixed = TRUE), ifelse(round(runif(n, 1, 2)) == 1, "North", "South"),
                                       NA)))))
  }
  
  # Create a vectorized version of the millRoom decision logic
  get_mill_room <- function(mix_room) {
    n <- length(mix_room)
    ifelse(mix_room == "North" | mix_room == "South", mix_room,
           ifelse(mix_room == "Old" | mix_room == "New", ifelse(round(runif(n, 1, 2)) == 1, "North", "South"),
                  NA))
  }
  
  # Apply the decision logic to the dataMaster data frame
  dataMaster$mixRoom <- get_mix_room(dataMaster$BatchSize)
  dataMaster$millRoom <- get_mill_room(dataMaster$mixRoom)
  
  # Generate run order and arrange the data frame by this order
  # set.seed(123) # For reproducibility, if desired
  dataMaster$runOrder <- sample(nrow(dataMaster))
  return(dataMaster[order(dataMaster$runOrder),])

  
  
}

prepData <- function(){
  
}

sub5000Sched <- function(df){
  df <- data[i,]
  size <- df$literSize
  mixingRoom <- df$mixRoom
  millingRoom <- df$millRoom
  cleanGroup <- df$Group
  data <- setDT(data)
  nextItemGroup <- data[mixRoom == mixingRoom & runOrder > i][1]$Group
  nextItemGroup <- ifelse(is.na(nextItemGroup), 'C',nextItemGroup)
  #selects the next prod being mixed and finds what cleaning group it belongs to.
  millingRoom <- df$millRoom
  nextMillGroup <- data[millRoom == millingRoom & runOrder > i][1]$Group
  nextMillGroup <- ifelse(is.na(nextMillGroup), 'C',nextMillGroup)
  # Get the millingRoom value from df
  itemNo <- df$`Production Order[Item No_]`
  prodNo <- df$`Production Order[No_]`
  # print(paste(mixingRoom, millingRoom, itemNo))
  setupTime <- df$`Setup Time`
  millingTime <- floor(.3 * df$`Mixing Time`)
  mixingTime <- ceiling(.7 * df$`Mixing Time`)
  dryCleanSetupTime <- df$DrySetupMins
  wetCleanSetupTime <- df$WetSetupMins
  dryCleanTime <- df$DryCleanMins
  wetCleanTime <- df$WetCleanMins
  dryCleanReset <- df$DryTearDownMins
  wetCleanReset <- df$WetTearDownMins
  mixingRoomTotalTime <- FALSE
  millingRoomTotalTime <- FALSE
  #vector Definition
  millRoomSetup <- c(1:setupTime)
  names(millRoomSetup) <- c(rep("Setup", setupTime))
  
  if (df$literSize != '5000') {
    if (cleanGroup == 'A' & nextItemGroup == cleanGroup) {
      totalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                     rep(paste("Milling", itemNo, prodNo, sep = " - "), millingTime),
                     rep(paste("Mixing", itemNo, prodNo, sep = " - "), mixingTime),
                     rep(paste("Dry Clean Setup", itemNo, prodNo, sep = " - "), dryCleanSetupTime),
                     rep(paste("Dry Clean", itemNo, prodNo, sep = " - "), dryCleanTime),
                     rep(paste("Dry Clean teardown", itemNo, prodNo, sep = " - "), dryCleanReset))
      return(totalTime)
    }
    if (cleanGroup != 'A' | nextItemGroup != cleanGroup) {
      totalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                     rep(paste("Milling", itemNo, prodNo, sep = " - "), millingTime), 
                     rep(paste("Mixing", itemNo, prodNo, sep = " - "), mixingTime),
                     rep(paste("Wet Clean Setup", itemNo, prodNo, sep = " - "), wetCleanSetupTime),
                     rep(paste("Wet Clean", itemNo, prodNo, sep = " - "), wetCleanTime),
                     rep(paste("Drying", itemNo, prodNo, sep = " - "), 90),
                     rep(paste("Wet Clean teardown", itemNo, prodNo, sep = " - "), wetCleanReset))
      return(totalTime)
    }
  }
}

sup5000Sched <- function(df){
  df <- data[i,]
  size <- df$literSize
  mixingRoom <- df$mixRoom
  millingRoom <- df$millRoom
  cleanGroup <- df$Group
  data <- setDT(data)
  nextItemGroup <- data[mixRoom == mixingRoom & runOrder > i][1]$Group
  nextItemGroup <- ifelse(is.na(nextItemGroup), 'C',nextItemGroup)
  #selects the next prod being mixed and finds what cleaning group it belongs to.
  millingRoom <- df$millRoom
  nextMillGroup <- data[millRoom == millingRoom & runOrder > i][1]$Group
  nextMillGroup <- ifelse(is.na(nextMillGroup), 'C',nextMillGroup)
  # Get the millingRoom value from df
  itemNo <- df$`Production Order[Item No_]`
  prodNo <- df$`Production Order[No_]`
  # print(paste(mixingRoom, millingRoom, itemNo))
  setupTime <- df$`Setup Time`
  millingTime <- floor(.3 * df$`Mixing Time`)
  mixingTime <- ceiling(.7 * df$`Mixing Time`)
  dryCleanSetupTime <- df$DrySetupMins
  wetCleanSetupTime <- df$WetSetupMins
  dryCleanTime <- df$DryCleanMins
  wetCleanTime <- df$WetCleanMins
  dryCleanReset <- df$DryTearDownMins
  wetCleanReset <- df$WetTearDownMins
  mixingRoomTotalTime <- FALSE
  millingRoomTotalTime <- FALSE
  #vector Definition
  millRoomSetup <- c(1:setupTime)
  names(millRoomSetup) <- c(rep("Setup", setupTime))
  
  #creates schedule of 5000L products

  if (cleanGroup == 'A' & nextItemGroup == 'A') {
    mixingRoomTotalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                             rep(paste("Waiting - Milling", itemNo, prodNo, sep = " - "), millingTime), 
                             rep(paste("Mixing", itemNo, prodNo, sep = " - "), mixingTime),
                             rep(paste("Dry Clean Setup", itemNo, prodNo, sep = " - "), dryCleanSetupTime),
                             rep(paste("Dry Clean", itemNo, prodNo, sep = " - "), dryCleanTime),
                             rep(paste("Dry Clean teardown", itemNo, prodNo, sep = " - "), dryCleanReset))
  }
  if (cleanGroup == 'A' & nextMillGroup == 'A') {
    millingRoomTotalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                              rep(paste("Milling", itemNo, prodNo, sep = " - "), millingTime), 
                              rep(paste("Dry Clean Setup", itemNo, prodNo, sep = " - "), dryCleanSetupTime),
                              rep(paste("Dry Clean", itemNo, prodNo, sep = " - "), dryCleanTime),
                              rep(paste("Dry Clean teardown", itemNo, prodNo, sep = " - "), dryCleanReset))
  }
  if (cleanGroup == 'C' | nextItemGroup != 'A') {
    mixingRoomTotalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                             rep(paste("Waiting - Milling", itemNo, prodNo, sep = " - "), millingTime),  
                             rep(paste("Mixing", itemNo, prodNo, sep = " - "), mixingTime),
                             rep(paste("Wet Clean Setup", itemNo, prodNo, sep = " - "), wetCleanSetupTime),
                             rep(paste("Wet Clean", itemNo, prodNo, sep = " - "), wetCleanTime),
                             rep(paste("Drying", itemNo, prodNo, sep = " - "), 90),
                             rep(paste("Wet Clean teardown", itemNo, prodNo, sep = " - "), wetCleanReset))
  }
  
  if (cleanGroup == 'C' | nextMillGroup != 'A') {
    millingRoomTotalTime <- c(rep(paste("Setup", itemNo, prodNo, sep = " - "), setupTime), 
                              rep(paste("Milling", itemNo, prodNo, sep = " - "), millingTime), 
                              rep(paste("Wet Clean Setup", itemNo, prodNo, sep = " - "), wetCleanSetupTime),
                              rep(paste("Wet Clean", itemNo, prodNo, sep = " - "), wetCleanTime),
                              rep(paste("Drying", itemNo, prodNo, sep = " - "), 90),
                              rep(paste("Wet Clean teardown", itemNo, prodNo, sep = " - "), wetCleanReset))
  }
  return(list(mixingRoomTotalTime = mixingRoomTotalTime, 
              millingRoomTotalTime = millingRoomTotalTime))
  
}

plot_schedule <- function(df) {
  
  # Convert time to datetime
  df <- df %>% 
    pivot_longer(!time, names_to = "Room", values_to = "Operation") %>%
    mutate(time = as.POSIXct(time, format="%H:%M"))
  
  # Group by Operation and room, and then summarize to get the start and end time of each Operation
  df_summarized <- df %>%
    group_by(Operation, Room) %>%
    summarize(start_time = min(time), end_time = max(time) + 60) %>%
    ungroup() %>%
    mutate(
      mid_time = start_time + (end_time - start_time) / 2,
      mix_room =
        case_when(
          Room == 'North' ~ 1,
          Room == 'South' ~ 2,
          Room == 'Old'   ~ 3,
          Room == 'New'   ~ 4
          
        )
    )
  
  # Plot
  plot <- ggplot(df_summarized, aes(xmin = start_time, xmax = end_time, ymin = mix_room - 0.3, ymax = mix_room + 0.3)) +
    geom_rect(aes(fill = Operation), color = "black", size = 0.5) +
    geom_text(aes(x = mid_time, y = mix_room, label = Operation), color = "white", fontface = "bold") +
    coord_flip() +
    labs(title = "Gantt Chart", x = "Time", y = "Room") +
    theme_minimal() +
    scale_x_datetime(date_breaks = "60 min", date_labels = "%H:%M") +
    theme(legend.position = "none") +
    scale_y_continuous(labels = unique(df$Room), breaks = 1:length(unique(df$Room)))
  
  return(list(plot, df_summarized))
}



data <- prepData()
bestBatchCount <- 0
bestIdleCount <- 0
bestRoomScedule <- create_empyty_room_Schedule()
nruns <- 1000
mustRun <- c('395428', '400008')


dataPrepTime <- 0
#total time spent assigning variables
variableTime <- 0
#total time spent creating schedule vectors
vectorTime <- 0
#total time spent creating schedule 
scheduleTime <- 0
#total time spent counting idles
idleCntTime <- 0
#total time spent evaluating sched
bestTime <- 0 

# startTime <- Sys.time()

for (ii in 1:nruns) {
  
  batchCount <- 0
  idleCount <- 0
  startTime <- Sys.time()
  data <- prepData()
  dataPrepTime <- dataPrepTime + Sys.time() - startTime
  roomSchedule <- create_empyty_room_Schedule()
  PO_completed <- c()

  for (i in 1:nrow(data)) {
    # Extract data for the current row once
    startTime <- Sys.time()
    df <- data[i,]
    size <- df$literSize
    mixingRoom <- df$mixRoom
    millingRoom <- df$millRoom
    cleanGroup <- df$Group
    data <- setDT(data)
    nextItemGroup <- data[mixRoom == mixingRoom & runOrder > i][1]$Group
    nextItemGroup <- ifelse(is.na(nextItemGroup), 'C',nextItemGroup)
    #selects the next prod being mixed and finds what cleaning group it belongs to.
    millingRoom <- df$millRoom
    nextMillGroup <- data[millRoom == millingRoom & runOrder > i][1]$Group
    nextMillGroup <- ifelse(is.na(nextMillGroup), 'C',nextMillGroup)
    # Get the millingRoom value from df
    itemNo <- df$`Production Order[Item No_]`
    prodNo <- df$`Production Order[No_]`
    # print(paste(mixingRoom, millingRoom, itemNo))
    setupTime <- df$`Setup Time`
    millingTime <- floor(.3 * df$`Mixing Time`)
    mixingTime <- ceiling(.7 * df$`Mixing Time`)
    dryCleanSetupTime <- df$DrySetupMins
    wetCleanSetupTime <- df$WetSetupMins
    dryCleanTime <- df$DryCleanMins
    wetCleanTime <- df$WetCleanMins
    dryCleanReset <- df$DryTearDownMins
    wetCleanReset <- df$WetTearDownMins
    mixingRoomTotalTime <- FALSE
    millingRoomTotalTime <- FALSE
    #vector Definition
    millRoomSetup <- c(1:setupTime)
    names(millRoomSetup) <- c(rep("Setup", setupTime))
    
    variableTime <- variableTime + Sys.time() - startTime
    
    #################################################
    
    startTime <- Sys.time()
    #creates schedule of products under 5000L
    if (df$literSize != '5000') {
      totalTime <- sub5000Sched(df)
    }
    
    #creates schedule of 5000L products
    if (df$literSize == '5000') {
      temp <- sup5000Sched(df)
      mixingRoomTotalTime <- unlist(temp[1])
      millingRoomTotalTime <- unlist(temp[2])
      
    }
    vectorTime <- vectorTime + Sys.time() - startTime
    
    ##################################################
    
    startTime <- Sys.time()
    if (df$literSize != '5000') {
      firstIdleTime <- which(roomSchedule[[mixingRoom]] == 'Idle')[1]
      
      if ((nrow(roomSchedule) - firstIdleTime + 1) > length(totalTime)) {
        runLength <- length(totalTime)
        roomSchedule[[mixingRoom]][firstIdleTime:(runLength + firstIdleTime - 1)] <- totalTime
        roomSchedule[[mixingRoom]][which(roomSchedule[[mixingRoom]][1:firstIdleTime] == "Idle")] <- paste("Idle", i)
        batchCount <- batchCount + 1
        PO_completed[batchCount] <- prodNo
      }
      
    }
    
    firstIdleTimeMixingRoom <- which(roomSchedule[[mixingRoom]] == 'Idle')[1]
    firstIdleTimeMillingRoom <- which(roomSchedule[[millingRoom]] == 'Idle')[1]
    if (firstIdleTimeMillingRoom > firstIdleTimeMixingRoom) {
      firstIdleTimeMixingRoom <- firstIdleTimeMillingRoom
    }
    
    if (df$literSize == '5000' & (nrow(roomSchedule) - firstIdleTimeMillingRoom) > length(millingRoomTotalTime) & (nrow(roomSchedule) - firstIdleTimeMixingRoom) > length(mixingRoomTotalTime)) {
      runLength <- length(millingRoomTotalTime)
      roomSchedule[[millingRoom]][firstIdleTimeMillingRoom:(firstIdleTimeMillingRoom + runLength - 1)] <- millingRoomTotalTime
      runLength <- length(mixingRoomTotalTime)
      roomSchedule[[mixingRoom]][firstIdleTimeMixingRoom:(firstIdleTimeMixingRoom + runLength - 1)] <- mixingRoomTotalTime
      roomSchedule[[mixingRoom]][which(roomSchedule[[mixingRoom]][1:firstIdleTimeMixingRoom] == "Idle")] <- paste("Idle", i)
      roomSchedule[[millingRoom]][which(roomSchedule[[millingRoom]][1:firstIdleTimeMillingRoom] == "Idle")] <- paste("Idle", i)
      batchCount <- batchCount + 1
      PO_completed[batchCount] <- prodNo
      
    }
    
    scheduleTime <- scheduleTime + Sys.time() - startTime
    
  }
  
  #######################################################
  
  startTime <- Sys.time()
  # idleCount <- sum(apply(roomSchedule, 1:2, function(x) grepl("Idle", x)))
  idleCount <- sum(grepl("Idle", unlist(roomSchedule[, 2:5])))
  
  
  idleCntTime <- idleCntTime + Sys.time() - startTime
  
  startTime <- Sys.time()
  if ((batchCount > bestBatchCount | 
       (batchCount >= bestBatchCount & idleCount > bestIdleCount)) &
      all(mustRun %in% PO_completed)) {
    
    print(paste("New Best Schedule Found. Iteration:", ii))
    print(paste("Batches ran:", batchCount))
    print(paste("Previous Batches ran:", bestBatchCount))
    print(paste("Idle Time:", idleCount))
    print(paste("Previous Idle Time:", bestIdleCount))
    print(PO_completed)
    
    bestBatchCount <- batchCount
    bestIdleCount <- idleCount
    bestRoomScedule <- roomSchedule
    
  }
  bestTime <- bestTime + Sys.time() - startTime
  
  ########################################################
}
# print(Sys.time() - startTime)
plot <- plot_schedule(bestRoomScedule)
plot
View(plot[[2]])

write.csv(plot[[2]], 'C:/Users/arose/OneDrive - Ortho Molecular Products/Desktop/Schedule_Optimization/abbrvRoomSchedule.csv')
write.csv(bestRoomScedule, 'C:/Users/arose/OneDrive - Ortho Molecular Products/Desktop/Schedule_Optimization/fullRoomSchedule.csv')


startTime <- Sys.time()

# data[['ItemNo_BatchQTY_key']]
# data$ItemNo_BatchQTY_key
data[,1]

print(Sys.time() - startTime)






