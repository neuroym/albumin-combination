#1 多因素cox回归
#1.1安装必要工具包
install.packages("tidyverse")
install.packages("gtsummary")
install.packages("tidyr")
install.packages("survey")
install.packages("survminer")
install.packages("flextable")
install.packages("survival")
install.packages("readxl")
install.packages("plyr")

#1.2加载必备工具包
library(tidyverse)
library(gtsummary)
library(tidyr)
library(survey)
library(survminer)
library(flextable)
library(survival)
library(readxl)
library(plyr)


#1.3导入数据
unicox <- read_excel("D:/albumin/cox.xlsx")
str(cox120)
head(cox)
rm(cox)

#1.4批量单因素回归
y<- Surv(time=cox120$time,event = cox120$events==1)
#1.42批量单因素回归模型建立
Uni_cox_model<- function(x){
  FML<- as.formula(paste("y~",x))
  cox<- coxph(FML,data=cox120)
  cox1<- summary(cox)
  HR<- round(cox1$coefficients[,2],2)
  Pvalue<- round(cox1$coefficients[,5],3)
  CI5<- round(cox1$conf.int[,3],2)
  CI95<- round(cox1$conf.int[,4],2)
  Uni_cox_model<- data.frame('characteristics'=x,
                             'HR'=HR,
                             'CI5'=CI5,
                             'CI95'=CI95,
                             'p'=Pvalue)
  return(Uni_cox_model)}
#1.43将想要的单因素回归变量输入模型
#1.431查看变量的名字和序号
names(unicox)
#1.432输入变量序号
variable.names<- colnames(cox120)[4:43]
#1.44输出结果
Uni_cox<- lapply(variable.names,Uni_cox_model)
Uni_cox<- ldply(Uni_cox,data.frame)
#1.45优化表格，将可信区间合并为一个单元格
Uni_cox$CI<-paste(Uni_cox$CI5,'-',Uni_cox$CI95)
Uni_cox<-Uni_cox[,-3:-4]
#1.46查看单因素结果
View(Uni_cox)
#1.47输出结果
install.packages("openxlsx")
library(openxlsx)
#1.471转变为数据框
coef_df <- as.data.frame(Uni_cox)
#1.472创建一个新的excel
wb <- createWorkbook()
#1.473  添加一个工作表
addWorksheet(wb, "Cox Regression Results")
#1.474写入数据到数据框
writeData(wb, sheet = 1, x = coef_df)
#1.475保存Excel文件到指定路径
file_path <- "D:/albumin/coxresult/Uni_cox120.xlsx"  # 替换为你指定的路径
saveWorkbook(wb, file = file_path, overwrite = TRUE)




#2多因素cox回归
#2.1 指定需要转换为因子的列名
factor_vars <- c("sex","group","combinetime","hypertension","congestive_heart_failure","respiratory_failure","coronary_atherosclerosis","hyperlipidemia","myocardial_infarct","atrial_fibrillation","diabetes","chronic_renal_disease"
                 ,"chronic_pulmonary_disease","rrt","ventilation","endovascular_therapy","has_vaso","has_tpa","has_alb")
# 批量转换这些列转换为factor
cox120[factor_vars] <- lapply(cox120[factor_vars], as.factor)
str(cox120)
view(cox)
rm(data120)


#1.4批量单因素回归
y<- Surv(time=cox$time,event = cox$events==1)
#1.42批量单因素回归模型建立
Uni_cox_model1<- function(x){
  FML<- as.formula(paste("y~",x))
  coxf<- coxph(FML,data=cox)
  coxf1<- summary(coxf)
  HR<- round(coxf1$coefficients[,2],2)
  Pvalue<- round(coxf1$coefficients[,5],3)
  CI5<- round(coxf1$conf.int[,3],2)
  CI95<- round(coxf1$conf.int[,4],2)
  Uni_cox_model1<- data.frame('characteristics'=x,
                             'HR'=HR,
                             'CI5'=CI5,
                             'CI95'=CI95,
                             'p'=Pvalue)
  return(Uni_cox_model1)}
#1.43将想要的单因素回归变量输入模型
#1.431查看变量的名字和序号
names(cox)
#1.432输入变量序号
variable.names<- colnames(cox)[4:41]
#1.44输出结果
f_cox<- lapply(variable.names,Uni_cox_model1)
f_cox<- ldply(f_cox,data.frame)
#1.45优化表格，将可信区间合并为一个单元格
f_cox$CI<-paste(f_cox$CI5,'-',f_cox$CI95)
f_cox<-f_cox[,-3:-4]
#1.46查看单因素结果
View(f_cox)
#1.47输出结果
library(openxlsx)
#1.471转变为数据框
coef_df <- as.data.frame(f_cox)
#1.472创建一个新的excel
wb <- createWorkbook()
#1.473  添加一个工作表
addWorksheet(wb, "Cox Regression Results")
#1.474写入数据到数据框
writeData(wb, sheet = 1, x = coef_df)
#1.475保存Excel文件到指定路径
file_path <- "D:/albumin/f_cox.xlsx"  # 替换为你指定的路径
saveWorkbook(wb, file = file_path, overwrite = TRUE)

#2.2构建cox回归模型
res.cox1201<-coxph(Surv(time,events) ~ group+age+sex,data=data120)
res
res.cox1<-coxph(Surv(time,events) ~ group+age+sex+first_day_wbc_max+apsiii+lactate,data=cox)
res.cox2<-coxph(Surv(time,events) ~ combinetime+age+sex+sofa_score+gcs_min+apsiii+hypertension+congestive_heart_failure+atrial_fibrillation+diabetes+chronic_renal_disease
                +chronic_pulmonary_disease+rrt+ventilation+endovascular_therapy+vent_time+has_vaso,data=unicox)
#查看统计结果
summary(res.cox)
summary(res.cox1)
summary(res.cox2)
#输出统计结果
m1_tal_cox<- tbl_regression(res.cox2,exponentiate = TRUE)
m1_tal_cox
#将结果导出为word
#1转换为flextbale
coxtable<- as_flex_table(m1_tal_cox)
#2导出word，其中 "D:/albumin/coxtable.docx"为你想导出的储存位置及文件名
save_as_docx(coxtable,path = "D:/albumin/coxtable.docx")

res.cox3<-coxph(Surv(time,events) ~ combinetime+age+sex+sofa_score+gcs_min+apsiii+hypertension+congestive_heart_failure+atrial_fibrillation+diabetes+chronic_renal_disease
                +chronic_pulmonary_disease+rrt+ventilation+endovascular_therapy+vent_time+has_vaso,data=cox)
summary(res.cox3)

# 假设数据框是df，生存时间变量是time，事件变量是events
variables <- names(cox)[!names(cox) %in% c("time", "events","has_cry","carcinoma")]
results <- lapply(variables, function(var) {
  fit <- coxph(as.formula(paste("Surv(time, events) ~", var)), data = cox)
 })
summary(results)
unicox <- tbl_regression(results,exponentiate = TRUE)
# 查看每个因子变量的水平数
sapply(cox[factor_vars], function(x) length(levels(x)))

#####另一种批量单因素cox回归
# 载入必要的包
library(survival)
library(dplyr)
library(readxl)
library(purrr)

# 读取数据
cox210 <- read_excel("your_data.xlsx") # 请根据实际文件路径修改

# 定义生存对象
y <- Surv(time = cox120$time, event = cox120$events == 1)

# 变量列表
variables <- colnames(cox120)[!colnames(cox120) %in% c("time", "events")]

# 单因素COX回归分析函数
uni_cox_model <- function(var) {
  formula <- as.formula(paste("y ~", var))
  cox_model <- coxph(formula, data = cox120)
  cox_summary <- summary(cox_model)
  
  # 提取结果
  HR <- round(cox_summary$coefficients[, "exp(coef)"], 2)
  CI5 <- round(cox_summary$conf.int[, "lower .95"], 2)
  CI95 <- round(cox_summary$conf.int[, "upper .95"], 2)
  Pvalue <- round(cox_summary$coefficients[, "Pr(>|z|)"], 3)
  
  result <- data.frame(
    characteristics = var,
    HR = HR,
    CI5 = CI5,
    CI95 = CI95,
    p = Pvalue
  )
  
  return(result)
}

# 批量执行
results <- map_dfr(variables, uni_cox_model)

# 查看结果
print(results)

#######COX亚组森林图
install.packages("jstable")
install.packages("devtools")
install.packages("cli")
install.packages("usethis")
install.packages("ellipsis")
install.packages("forestploter")

library(survival)
library(jstable)
library(cli)
library(usethis)
library(devtools)
library(grid)
library(forestploter)
##分类变量转化为因子变量，并设置标签
suppressMessages(library(tidyverse))
datasub<- cox120 %>%
  mutate(group=as.factor(group)) %>%
  select(time,status,group,sex,admission_age,mbp,resp_rate,first_day_wbc_max,first_day_bun_max,hypertension,first_day_creatinine_max,first_day_hemoglobin_max,
         lactate,sofa_score,apsiii,congestive_heart_failure,diabetes,respiratory_failure,chronic_renal_disease,chronic_pulmonary_disease,atrial_fibrillation,
         hyperlipidemia,rrt) %>% 
  mutate(sex=factor(sex,levels=c(0,1),labels=c("Female","Male")), admission_age=ifelse(admission_age >65,">65","<=65"),
         admission_age=factor(admission_age, levels=c(">65","<=65")),
         mbp=ifelse(mbp >80.15,">80.15","<=80.15"),
         mbp=factor(mbp, levels=c(">80.15","<=80.15")),
         resp_rate=ifelse(resp_rate >18.67,">18.67","<=18.67"),
         resp_rate=factor(resp_rate, levels=c(">18.67","<=18.67")),
         first_day_wbc_max=ifelse(first_day_wbc_max >14.50,">14.50","<=14.50"),
         first_day_wbc_max=factor(first_day_wbc_max, levels=c(">14.50","<=14.50")),
         first_day_bun_max=ifelse(first_day_bun_max >19,">19","<=19"),
         first_day_bun_max=factor(first_day_bun_max, levels=c(">19","<=19")),
         first_day_creatinine_max=ifelse(first_day_creatinine_max >1,">1","<=1"),
         first_day_creatinine_max=factor(first_day_creatinine_max, levels=c(">1","<=1")),
         first_day_hemoglobin_max=ifelse(first_day_hemoglobin_max >11.70,">11.70","<=11.70"),
         first_day_hemoglobin_max=factor(first_day_hemoglobin_max, levels=c(">11.70","<=11.70")),
         lactate=ifelse(lactate >2.90,">2.90","<=2.90"),
         lactate=factor(lactate, levels=c(">2.90","<=2.90")),
         sofa_score=ifelse(sofa_score >3,">3","<=3"),
         sofa_score=factor(sofa_score, levels=c(">3","<=3")),
         apsiii=ifelse(apsiii >42,">42","<=42"),
         apsiii=factor(apsiii, levels=c(">42","<=42")),
         hypertension=factor(hypertension, levels=c(0,1),labels=c("No","Yes")),
         congestive_heart_failure=factor(congestive_heart_failure, levels=c(0,1),labels=c("No","Yes")),
         diabetes=factor(diabetes, levels=c(0,1),labels=c("No","Yes")),
         respiratory_failure=factor(respiratory_failure, levels=c(0,1),labels=c("No","Yes")),
         chronic_renal_disease=factor(chronic_renal_disease, levels=c(0,1),labels=c("No","YES")),
         chronic_pulmonary_disease=factor(chronic_pulmonary_disease, levels=c(0,1),labels=c("No","YES")),
         hyperlipidemia=factor(hyperlipidemia, levels=c(0,1),labels=c("No","Yes")),
         atrial_fibrillation=factor(atrial_fibrillation, levels=c(0,1),labels=c("No","Yes")),
         rrt=factor(rrt, levels=c(0,1),labels=c("No","Yes")),
         group=factor(group, levels=c(0,1),labels=c("Crystalloids alone","Combination"))
  )

##cox亚组分析和P交互值
res1<-TableSubgroupMultiCox(formula=Surv(time,status)~group,var_subgroup=c("sex","admission_age","mbp","resp_rate","first_day_wbc_max","first_day_bun_max","hypertension","first_day_creatinine_max","first_day_hemoglobin_max",
               "lactate","sofa_score","apsiii","congestive_heart_failure","diabetes","respiratory_failure","chronic_renal_disease","chronic_pulmonary_disease","atrial_fibrillation",
               "hyperlipidemia","rrt"),
data=datasub
)
res1
view(res1)
warning(res)
pbc<-pbc
attach(pbc)
view(pbc)
###将数据框中events重命名为status
library(dplyr)

cox120 <- cox120 %>%
  dplyr::rename(status = events)
view(cox120)

suppressMessages(library(tidyverse))
datasub1<- cox210 %>%
  mutate(group=as.factor(group)) %>%
  select(time,status,group,sex,admission_age,first_day_wbc_max,lactate,apsiii) %>% 
  mutate(sex=factor(sex,levels=c(0,1),labels=c("Female","Male")), admission_age=ifelse(admission_age >65,">65","<=65"),
         admission_age=factor(admission_age, levels=c(">65","<=65")),
         first_day_wbc_max=ifelse(first_day_wbc_max >14.50,">14.50","<=14.50"),
         first_day_wbc_max=factor(first_day_wbc_max, levels=c(">14.50","<=14.50")),
         lactate=ifelse(lactate >2.90,">2.90","<=2.90"),
         lactate=factor(lactate, levels=c(">2.90","<=2.90")),
         apsiii=ifelse(apsiii >42,">42","<=42"),
         apsiii=factor(apsiii, levels=c(">42","<=42")),
         group=factor(group, levels=c(0,1),labels=c("Crystalloids alone","Combination"))
  )

##cox亚组分析和P交互值
res<-TableSubgroupMultiCox(formula=Surv(time,status)~group,
    var_subgroup=c("sex","admission_age","first_day_wbc_max","lactate","apsiii"),
    data=datasub1
)
###整理数据
write.csv(res, file = "D:/albumin/res.csv", row.names = FALSE) #导出数据，便于自己调整标题等信息
rm(res)
view(plot_cox210)
str(plot_cox210)
plot_cox210<- res #对res重命名
plot_cox210[,c(2,3,9,10)][is.na(plot_cox210[,c(2,3,9,10)])] <- ""  ##选取2，3，9，10列数据在森林图中显示，并替换缺失值NA为一个空格符号
plot_cox210$''<- paste(rep("",nrow(plot_cox210)),collapse="") #添加空白列，用于存放森林图的图形部分
plot_cox210[,4:6]<-apply(plot_cox210[,4:6],2,as.numeric) #将4到6列数据转换为数值型
plot_cox210$"HR(95%CI)"<- ifelse(is.na(plot_cox210$"Point Estimate"),"",sprintf("%.2f(%.2f to %.2f)",plot_cox210$"Point Estimate",
                                                                                plot_cox210$Lower,
                                                                                plot_cox210$Upper)) ##计算95%CI，以便显示在图形中
plot_cox210 #查看数据

###绘图
plot_cox210$Cry <- ifelse(is.na(plot_cox210$Cry),"",plot_cox210$Cry)
plot_cox210$Comb <- ifelse(is.na(plot_cox210$Comb),"",plot_cox210$Comb)
plot_cox210$`P for interaction`<- ifelse(is.na(plot_cox210$`P for interaction`),"",plot_cox210$`P for interaction`)


p210<- forest(
  data=plot_cox210[,c(1,7,8,11,12,10)], #选择需要用于绘图的列
  lower = plot_cox210$Lower,#置信区间下限
  upper = plot_cox210$Upper,#置信区间上限
  est=plot_cox210$'Point Estimate', #点估计值
  ci_column=4, #点估计对应的列
  ref_line=1, #设置参考线位置
  sizes = (plot_cox210$`Point Estimate`+0.001)*0.3,
  xlim=c(0,4)#X轴的范围
)
plot(p210)
