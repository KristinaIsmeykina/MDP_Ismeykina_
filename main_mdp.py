import function_mdp
import win32com.client 
import pandas as pd
import numpy as np
import sys 

rastr = win32com.client.Dispatch("Astra.Rastr")
regime = r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\regime.rg2'
shablon_rgm = "C:/Users/Кристина/Documents/RastrWin3/SHABLON/режим.rg2"
shablon_ut = "C:/Users/Кристина/Documents/RastrWin3/SHABLON/траектория утяжеления.ut2"
shablon_sech = "C:/Users/Кристина/Documents/RastrWin3/SHABLON/сечения.sch"
rastr.Load(1, regime, shablon_rgm) 
rastr.NewFile(shablon_ut)
rastr.NewFile(shablon_sech)

#формирование траектрии утяжеления
tra_ut = pd.read_csv(r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\vector.csv')
function_mdp.set_tra_ut(rastr, tra_ut)
rastr.Save(r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\траектория утяжеления.ut2', shablon_ut)

#формирование сечения
sechen = pd.read_json(r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\flowgate.json')
sechen = sechen.T #переставление индексов и колонок
sechen.index=np.arange(sechen.shape[0]) #замена индексации
function_mdp.set_sechen(rastr, sechen)
rastr.Save(r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\сечения.sch',shablon_sech)
current_flow = rastr.Tables('sechen').Cols('psech').Z(0)
print(print("Начальный переток в КС-", round(current_flow, 2)))
#утяжеление
function_mdp.do_ut(rastr)
limit_flow = abs(rastr.Tables('sechen').Cols('psech').Z(0))
print("Предельный переток в КС-", round(limit_flow, 2))

print(" ------------- 1 критерий по P в нормальном режиме.---------------  ")
# 1 критерий по P в нормальном режиме.
nereg = 30
mdp1 = 0.8*limit_flow - nereg
print("МДП по 1 критерию", round(mdp1,2))

print("---- проверочный критерий по U в нормальном режиме.--------------- ")
# проверочный критерий по U в нормальном режиме.
rastr.Load(1, regime, shablon_rgm) 
function_mdp.criterion_U_norm(rastr, 1.15)
mdpU = abs(rastr.Tables('sechen').Cols('psech').Z(0)) - nereg
print("Проверочный по U-", round(mdpU,2))

# 2 критерий по P в ПАР.
print("----------------------- 2 критерий по P в ПАР.-------------------- ")
rastr.Load(1, regime, shablon_rgm) 
faults = pd.read_json(r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\faults.json')
doavar_flow = function_mdp.criterion_P_par(rastr, faults, shablon_rgm)
print("Доаварийный переток в КС-", doavar_flow)
mdp2 = min(doavar_flow) - nereg
print("МДП по 2 критерию-", mdp2 )

# 3 критерий по U в ПАР
print(" -----------------------3 критерий по U в ПАР.--------------------- ")
rastr.Load(1, regime, shablon_rgm)
doavar_flow2 = function_mdp.criterion_U_par(rastr, faults, shablon_rgm)
mdp3 = min(doavar_flow2) - nereg
print("Доаварийный переток U-", doavar_flow2)
print("МДП по 3 критерию -", mdp3 )

# проверочный критерий по I в нормальном режиме
print(" ------------ проверочный критерий по I в нормальном режиме --------")
rastr.Load(1, regime, shablon_rgm)
function_mdp.criterion_I_norm(rastr,'i_dop_r', shablon_rgm)
mdpI = abs(rastr.Tables('sechen').Cols('psech').Z(0)) - nereg
print("Проверочный по I-", round(mdpI,2))

# 4 критерий по I в ПАР.
print(" ----------------------- 4 критерий по I в ПАР.-------------------- ")
rastr.Load(1, regime, shablon_rgm)

doavar_flow3 = function_mdp.criterion_I_par(rastr, faults, shablon_rgm)
mdp4 = min(doavar_flow3) - nereg
print("Доаварийный переток I-", doavar_flow3)
print("МДП по 4 критерию -", mdp4)

print("-----------------------РЕЗУЛЬТИРУЮЩИЙ МДП------------------------")
print(round(min(mdp1, mdp2, mdp3, mdp4, mdpU, mdpI), 2))