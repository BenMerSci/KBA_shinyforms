library(reshape2)
library(ggplot2)

tasks <- c("Review literature", "Mung data", "Stats analysis", "Write Report")
dfr <- data.frame(
  name        = factor(tasks, levels = tasks),
  start.date  = as.Date(c("2010-08-24", "2010-10-01", "2010-11-01", "2011-02-14")),
  end.date    = as.Date(c("2010-10-31", "2010-12-14", "2011-02-28", "2011-04-30")),
  is.critical = c(TRUE, FALSE, FALSE, TRUE)
)
mdfr <- melt(dfr, measure.vars = c("start.date", "end.date"))

ggplot(mdfr, aes(value, name, colour = is.critical)) + 
  geom_line(size = 6) +
  xlab(NULL) + 
  ylab(NULL)

