
regime = r'C:\Users\Кристина\Desktop\Нирс 1 семестр\3 семестр\Нирс\MDP_Ismeykina_\regime.rg2'

def set_tra_ut(rastr, tra_ut): 
    """This function sets the weighting vector
    Parameters:
    rastr (rastr): COM object
    tra_ut (dataframe): weighting vector
    """
    for row in range(tra_ut.shape[0]):
        rastr.Tables('ut_node').AddRow()  
        rastr.Tables('ut_node').Cols('ny').SetZ(row, tra_ut.loc[row,'node'])
        if (tra_ut.loc[row,'variable']=='pn' ):
            rastr.Tables('ut_node').Cols('pn').SetZ(row, tra_ut.loc[row,'value'])
            rastr.Tables('ut_node').Cols('tg').SetZ(row, tra_ut.loc[row,'tg'])
        elif (tra_ut.loc[row,'variable']=='pg'):
            rastr.Tables('ut_node').Cols('pg').SetZ(row, tra_ut.loc[row,'value'])
            rastr.Tables('ut_node').Cols('tip').SetZ(row, 2)  

def set_sechen(rastr, sechen): 
    """This function sets the flowgate
    Parameters:
    rastr (rastr): COM object
    sechen (dataFrame): weighting vector
    """
    rastr.Tables('sechen').AddRow()
    rastr.Tables('sechen').Cols('ns').SetZ(0, 1) #  № сечения 1
    for row in range(sechen.shape[0]):
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(row, 1)
        rastr.Tables('grline').Cols('ip').SetZ(row, sechen.loc[row,'ip'])
        rastr.Tables('grline').Cols('iq').SetZ(row, sechen.loc[row,'iq']) 

def do_ut(rastr): 
    """This function makes the mode heavier to obtain the maximum overflow
    Parameters:
    rastr (rastr): COM object
    """
    rastr.rgm('p')
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')

def faults_number(rastr, faults, shablon_rgm):
    """This function disconnects the branch according to the faults file
    Parameters:
    rastr (rastr): COM object
    faults[key] (string): name of fault
    shablon_rgm (string): template for creating a Rastrwin3 file
    
    Return:
    j (int): line to be disconnected
    """
    rastr.Load(1, regime, shablon_rgm)
    vetv = rastr.Tables('vetv')
    for j in range(0, vetv.size):
        if (faults['ip'] == vetv.Cols('ip').Z(j)) and \
       (faults['iq'] == vetv.Cols('iq').Z(j) and \
       faults['np']==vetv.Cols('np').Z(j)):
            return j

def criterion_U_norm(rastr, percent):
    """This function makes the mode heavier until the voltages go beyond the set limits.
    Parameters:
    rastr (rastr): COM object
    percent (int): Sets the stock percentage
    """
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
        if node.Cols('pn').Z(i) != 0:   #в узлах нагрузки напряжение не ниже percent*Uкр
            node.Cols('contr_v').SetZ(i, 1)
            nom = node.Cols('uhom').Z(i)
            node.Cols('umin').SetZ(i, nom*0.7*percent)
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')

def criterion_P_par(rastr, faults, shablon_rgm):
    """Second criterion
    Parameters:
    rastr (rastr): COM object
    faults (dataframe): faults
    shablon_rgm (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow
   """
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    #Отключим контроль P,U,I при утяжелении 
    enable_contr_set = rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, 0)
    vetv = rastr.Tables('vetv')
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] 
    for key in faults:
        fault = faults_number(rastr, faults[key], shablon_rgm)
        vetv.Cols('sta').SetZ(fault, faults[key]['sta'])
        rastr.rgm('p')
        if rastr.ut_utr('i') > 0:
            rastr.ut_utr('')
        limit_flow2 = abs(rastr.Tables('sechen').Cols('psech').Z(0))
        rastr.Load(1, regime, shablon_rgm)
        vetv.Cols('sta').SetZ(fault, faults[key]['sta'])
        rastr.rgm('')
        kd = rastr.step_ut("index")
        while (kd == 0) and abs(sechen.Cols('psech').Z(0)) < 0.92*limit_flow2:
            kd = rastr.step_ut("z")
            
        vetv.Cols('sta').SetZ(fault, 1-faults[key]['sta'])
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow

def criterion_U_par(rastr, faults, shablon_rgm):
    """Third criterion
    Parameters:
    rastr (rastr): COM object
    faults (dict): faults
    shablon_rgm (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow
    """   
    vetv = rastr.Tables('vetv')
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200) 
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] 
    for key in faults:
        fault = faults_number(rastr, faults[key],shablon_rgm)
        vetv.Cols('sta').SetZ(fault, faults[key]['sta'])
        criterion_U_norm(rastr, 1.1)
        vetv.Cols('sta').SetZ(fault, 1-faults[key]['sta'])
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow

def criterion_I_norm(rastr, i_dop):
    """This function makes the mode heavier until the currents go beyond the set limits.
    Parameters:
    rastr (rastr): COM object
    i_dop (string): Column name with specified current limit
    """
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
    
def criterion_I_par(rastr, faults, shablon_rgm):
    """Fourth criterion
    Parameters:
    rastr (rastr): COM object
    faults (dataframe): faults
    shablon_rgm (string): shablon
    
    Return:
    doavar_overflow (list): list of overflow
    """     
    vetv = rastr.Tables('vetv')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    sechen = rastr.Tables('sechen')
    doavar_flow = [] 
    for key in faults:
        fault = faults_number(rastr, faults[key],shablon_rgm)
        vetv.Cols('sta').SetZ(fault, faults[key]['sta'])
        rastr.rgm('p')
        
        #Перенос в расчетной модели Iадтн_рас--> Iддтн, утяжеление
        criterion_I_norm(rastr,'i_dop_r_av', shablon_rgm)
        vetv.Cols('sta').SetZ(fault, 1-faults[key]['sta'])
        rastr.rgm('p')
        doavar_flow.append(abs(sechen.Cols('psech').Z(0)))
    doavar_flow=[round(v, 2) for v in doavar_flow]
    return doavar_flow


    