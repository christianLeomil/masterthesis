#region superclasses


class Generator:
    def __init__(self,type_, id,eff,E_in,op_cost,inv_cost,emission):
        self.type_ = type_
        self.id = id
        self.eff = eff
        self.E_in= E_in
        self.inv_cost = inv_cost
        self.op_cost = op_cost
        self.emission = emission

class pv(Generator):
    def __init__(self,type_,id_number,eff,E_in,op_cost,inv_cost,emission):
        super().__init__('pv',id_number,eff,E_in,op_cost,inv_cost,emission)

    def libary():
        return ['P_pv','P_solar','pv_eff']
    # def generation_rule(model,t):
    #     return model.P_Pv_1[t] == model.P_solar[t] + model.pv_eff
    
    def generation_rule():
        return 'model.P_pv[t] == model.P_solar[t] * model.pv_eff'
    
class bat:
    def libary():
        return ['bat_SOC','bat_starting_SOC','P_bat_ch','P_bat_dis','bat_ch_eff','bat_dis_eff','E_bat_max']

    def battery_rule():
        string = """if t ==1: 
             return model.bat_SOC[t] == model.bat_starting_SOC
        else:
            return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_bat_ch[t-1] * model.bat_ch_eff -
                                                      model.P_bat_dis[t-1]/model.bat_dis_eff) * model.time_step / model.E_bat_max"""
        return string
    

class Testing:
    def demand_rule(model,t):
        return model.demand[t] == model.quant_x[t] + model.quant_y[t]

class Subclass(Testing):
    def __init__(self,model,t):
        super().__init__(model,t)

    def second_rule(model,t):
        return model.quant_x[t] == model.quant_z[t]
         
# endregion
# ---------------------------------------------------------------------------------------------------------------------