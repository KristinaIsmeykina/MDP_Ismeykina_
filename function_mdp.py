
regime = r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\regime.rg2'

def set_tra_ut(rastr, tra_ut): # формирование траектории утяжеления
    
    for row in range(tra_ut.shape[0]):
        rastr.Tables('ut_node').AddRow()  
        rastr.Tables('ut_node').Cols('ny').SetZ(row, tra_ut.loc[row,'node'])
        if (tra_ut.loc[row,'variable']=='pn' ):
            rastr.Tables('ut_node').Cols('pn').SetZ(row, tra_ut.loc[row,'value'])
            rastr.Tables('ut_node').Cols('tg').SetZ(row, tra_ut.loc[row,'tg'])
        elif (tra_ut.loc[row,'variable']=='pg'):
            rastr.Tables('ut_node').Cols('pg').SetZ(row, tra_ut.loc[row,'value'])
            rastr.Tables('ut_node').Cols('tip').SetZ(row, 2)  

def set_sechen(rastr, sechen): # формирование сечения
    rastr.Tables('sechen').AddRow()
    rastr.Tables('sechen').Cols('ns').SetZ(0, 1) #  № сечения 1
    for row in range(sechen.shape[0]):
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(row, 1)
        rastr.Tables('grline').Cols('ip').SetZ(row, sechen.loc[row,'ip'])
        rastr.Tables('grline').Cols('iq').SetZ(row, sechen.loc[row,'iq']) 

def do_ut(rastr): # утяжеление
    rastr.rgm('p')
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')

def crit_U_norm(rastr, percent):
    ut_common = rastr.Tables('ut_common')
    ut_common.Cols('iter').SetZ(0, 200)
    rastr.rgm('p')
    #Включим контроль U при утяжелении 
    #Включим контроль P,U,I при утяжелении 
    ut_common.Cols('enable_contr').SetZ(0, 1)
    #Отключим контроль I 
    ut_common.Cols('dis_i_contr').SetZ(0, 1)
    #Отключим контроль P 
    ut_common.Cols('dis_p_contr').SetZ(0, 1)
    #Включим контроль U
    ut_common.Cols('dis_v_contr').SetZ(0, 0)
    node = rastr.Tables('node') 
    for i in range(node.size):
        if node.Cols('pn').Z(i) != 0:   #в узлах нагрузки напряжение не ниже 1,15*Uкр
            node.Cols('contr_v').SetZ(i, 1)
            nom = node.Cols('uhom').Z(i)
            node.Cols('umin').SetZ(i, nom*0.7*percent)
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')

def crit_P_par(rastr, faults, shablon_rgm):
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    #Отключим контроль P,U,I при утяжелении 
    enable_contr_set = rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, 0)
    vetv = rastr.Tables('vetv')
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] # результаты для разных отключений
    for row in range(faults.shape[0]):
        memory = 0
        i = 0 
         # Отключаем линию
        while i < vetv.Size:      
            if ip.Z(i) == faults.loc[row,'ip'] and iq.Z(i) == faults.loc[row,'iq'] and np.Z(i) == faults.loc[row,'np']:
                vetv.Cols('sta').SetZ(i, 1)
                #memory = i
            i += 1
        rastr.rgm('p')
        if rastr.ut_utr('i') > 0:
            rastr.ut_utr('')
        limit_flow2 = abs(rastr.Tables('sechen').Cols('psech').Z(0))
        rastr.Load(1, regime, shablon_rgm)
        
        # Отключаем линию
        #vetv.Cols('sta').SetZ(memory, 1) # не удалось использовать memory
        while i < vetv.Size:      
            if ip.Z(i) == faults.loc[row,'ip'] and iq.Z(i) == faults.loc[row,'iq'] and np.Z(i) == faults.loc[row,'np']:
                vetv.Cols('sta').SetZ(i, 1)
            i += 1
        rastr.rgm('')
        kd = rastr.step_ut("index")
        while (kd == 0) and abs(sechen.Cols('psech').Z(0)) < 0.92*limit_flow2:
            kd = rastr.step_ut("z")
            
        # Включаем линию 
        #vetv.Cols('sta').SetZ(memory, 0) # не удалось использовать memory
        while i < vetv.Size:      
            if ip.Z(i) == faults.loc[row,'ip'] and iq.Z(i) == faults.loc[row,'iq'] and np.Z(i) == faults.loc[row,'np']:
                vetv.Cols('sta').SetZ(i, 0)
            i += 1
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow

def crit_U_par(rastr, faults, shablon_rgm):
    vetv = rastr.Tables('vetv')
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] # результаты для разных отключений
    for row in range(faults.shape[0]):
        memory = 0
        i = 0 
        # Отключаем линию
        while i < vetv.Size:      
            if ip.Z(i) == faults.loc[row,'ip'] and iq.Z(i) == faults.loc[row,'iq'] and np.Z(i) == faults.loc[row,'np']:
                vetv.Cols('sta').SetZ(i, 1)
                memory = i
            i += 1
        crit_U_norm(rastr, 1.1)
        # Включаем линию 
        vetv.Cols('sta').SetZ(memory, 0)
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow

def crit_I_norm(rastr, i_dop, shablon_rgm):
    ut_common = rastr.Tables('ut_common')
    ut_common.Cols('iter').SetZ(0, 200) 
    #Включим контроль I при утяжелении
    #Включим контроль P,U,I при утяжелении 
    ut_common.Cols('enable_contr').SetZ(0, 1)
    #Отключим контроль U 
    ut_common.Cols('dis_v_contr').SetZ(0, 1)
    #Отключим контроль P 
    ut_common.Cols('dis_p_contr').SetZ(0, 1)   
    #Включим контроль I 
    ut_common.Cols('dis_i_contr').SetZ(0, 0)  
    rastr.rgm('p')
    vetv = rastr.Tables('vetv')
    #Перенос в расчетной модели Iддтн_рас--> Iддтн
    i = 0 
    while i < vetv.Size:      
        vetv.Cols('i_dop').SetZ(i, vetv.Cols(i_dop).Z(i))     
        if vetv.Cols('i_dop').Z(i) != 0:
            vetv.Cols('contr_i').SetZ(i, 1 )
        i += 1
    
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')
    
def crit_I_par(rastr, faults, shablon_rgm):
          
    vetv = rastr.Tables('vetv')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] # результаты для разных отключений
    for row in range(faults.shape[0]):
        memory = 0
        i = 0 
        # Отключаем линию
        while i < vetv.Size:      
            if ip.Z(i) == faults.loc[row,'ip'] and iq.Z(i) == faults.loc[row,'iq'] and np.Z(i) == faults.loc[row,'np']:
                vetv.Cols('sta').SetZ(i, 1)
                memory = i
            i += 1
        rastr.rgm('p')
        
        #Перенос в расчетной модели Iадтн_рас--> Iддтн, утяжеление
        crit_I_norm(rastr,'i_dop_r_av', shablon_rgm)
        # Включаем линию 
        vetv.Cols('sta').SetZ(memory, 0)
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow


    