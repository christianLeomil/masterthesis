#region superclasses


class Generator:
    def __init__(self,type_, id,eff,E_in,op_cost,inv_cost,emission):
        self.class_type = 'generator'
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

    def solar_rule(model,t,n):
        return model.P_solar[t] * model.pv_eff[n] == model.P_from_pv[t,n]
    
class bat:
    def battery_rule(model,t,m):
            if t ==1:
                return model.SOC[t,m] == model.starting_SOC[m]
            else:
                return model.SOC[t,m] == model.SOC[t-1,m] + (model.P_to_bat[t-1,m] - model.P_from_bat[t-1,m]) * model.time_step / model.E_bat_max[m]

    def charge_limit(model,t,m):
        return model.P_to_bat[t,m] <= model.E_bat_max[m] * model.K_ch[t,m]
    
    def discharge_limit(model,t,m):
        return model.P_from_bat[t,m] <= model.E_bat_max[m] * model.K_dis[t,m]
    
    def keys_rule(model,t,m):
        return model.K_ch[t,m] + model.K_dis[t,m] <= 1
    
class general:
    def sell_energy(model,t):
        return model.E_sell[t] == model.P_to_net[t] * model.time_step
    
    def buy_energy(model,t):
        return model.E_buy[t] == model.P_net_demand[t] * model.time_step

# endregion
# ---------------------------------------------------------------------------------------------------------------------