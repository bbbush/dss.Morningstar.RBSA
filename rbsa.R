# Copyright 2015 Morningstar, Inc.

#library(xlsx)
#library(gdata)
library(XLConnect)
library(ggplot2)

# RBSA methodology https://corporate.morningstar.com/ib/documents/MethodologyDocuments/IBBAssociates/ReturnsBasedAnalysis.pdf

# data dump from PresentationStudio
ds<-(function(){
  xlsFile<-"OE Equity 1 - RBSA_20141010.xls"
  wb<-loadWorkbook(xlsFile)
  r=12
  p=26
  endCol = 2+1*8+(p-1)*8
  names<-readWorksheet(wb, 1, startRow=9, endRow=9, startCol=1, endCol=endCol, header=FALSE)
  names<-c(names[1:8], "period")
  months<-readWorksheet(wb, 1, startRow=7, endRow=8, startCol=1, endCol=endCol, header=FALSE)
  months<-paste(months[1,], months[2,], sep="~")[c(1,(1:(p-1))*8+2)]
  data<-readWorksheet(wb, 1, startRow=12, endRow=12+r-1, startCol=1, endCol=endCol, header=FALSE)
  ds<-data.frame()
  for(i in 1:p){
    t<-data[,c(1, 2, 1:6+ifelse(i==1, 2, 2+1+(i-1)*8))]
    t$period<-months[i]
    names(t)<-names
    ds<-rbind(ds, t)
  }
  if(any(round(rowSums(ds[,3:8])) != 100)){
    stop("input should be rescaled")
  }
  ds
})()

# The goal is to show style index and subject security on the same chart,
# also to show style drift. There is no clear meaning of X or Y axis, unlike
# the HB style charts. The chart simply assigns 2-D positions to represent
# the centroids, or center of masses. When the security is white noise data,
# cannot be explained by any of the style indexes, the center of mass should
# be (0, 0). As long as this criteria is met, the style index can be arranged
# in any position. The position of a security is also its center of mass, thus
# it is a weighted average of all masses of style indexes, with the weight
# provided in the dataset. When the assigned-position of style index changed,
# the chart may show different shape.

# 4, 6 or 9. In type "9", there is one style index whose assigned position is
# (0, 0), which means it has no mass and does not affect the shape of chart
matrix<-(makeMatrix<-function(type=6){
  switch(as.character(type),
         "4" = merge(c(-1, 1), c(1, -1)),
         "6" = merge(c(-1, 1), c(1, 0, -1)),
         "9" = merge(c(-1, 0, 1), c(1, 0, -1)))
})()

# points
xy<-(function(){
  xy<-as.data.frame(as.matrix(ds[,3:8])%*%as.matrix(matrix))
  cbind(SecId=ds$SecId, Name=ds[,2], Period=ds$period, as.data.frame(xy/100))
})()

# plot
(function(){
  rbsa<-(function(){
    r=1.5
    nameColors=sub("^Benchmark(.*?):\\s*", "", xy$Name)
    periodSizes=c(min(levels(xy$Period)), max(levels(xy$Period)))
    ggplot(xy)+
      geom_point(aes(x, y,
                     size=Period,
                     shape=grepl("^Benchmark", Name),
                     color=sub("^Benchmark(.*?):\\s*", "", Name)),
                 alpha=0.4)+
      scale_color_discrete(name="Security",
                           breaks=nameColors)+
      scale_shape_discrete(name="Benchmark")+
      scale_size_discrete(breaks=periodSizes)+
      guides(color = guide_legend(override.aes = list(size=3))) +
      guides(shape = guide_legend(override.aes = list(size=3))) +
      xlab("")+
      ylab("")+
      coord_fixed()+
      xlim(-r, r)+
      ylim(-r,r)+
      theme(legend.key = element_rect(fill="#ffffff"))+
      geom_point(data=matrix,
                 aes(matrix[1], matrix[2]),
                 size=3,
                 color="#000000",
                 alpha=0.3)+
      geom_text(data=matrix,
                aes(matrix[1], matrix[2], label=names(ds)[3:8]),
                size=4,
                vjust=1.2)
  })()
  
  # plot to file
  print(rbsa)
  png("rbsa.PNG", width=800, height=600)
  print(rbsa)
  dev.off()
})()
