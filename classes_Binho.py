import pyomo.environ as pyo
#region superclasses


class Generator:
<<<<<<< HEAD
    def __init__(self) -> None:
        self.new_method()

    def new_method(self):
        pass
=======
    def __init__(self,id) -> None:
        self.id = id
>>>>>>> b84b25036dd974f68e72fe27e1783896b92019db
    # def __init__(self,type_,model, id,eff,E_in,op_cost,inv_cost,emission):
    # # def __init__(self, model, id):
    #     self.class_type = 'generator'
    #     self.type_ = type_
    #     self.model = model
    #     self.id = id
    #     self.eff = eff
    #     self.E_in= E_in
    #     self.inv_cost = inv_cost
    #     self.op_cost = op_cost
    #     self.emission = emission
    

class Pv(Generator):
<<<<<<< HEAD
    # type = 'pv'
=======
    type = 'pv'
>>>>>>> b84b25036dd974f68e72fe27e1783896b92019db
    # def __init__(self,eff,E_in,op_cost,inv_cost,emission):
    #     # super().__init__('pv',id_number,eff,E_in,op_cost,inv_cost,emission)
    # # def __init__(self):
    #     self.eff = eff
    #     self.E_in= E_in
    #     self.inv_cost = inv_cost
    #     self.op_cost = op_cost
    #     self.emission = emission
<<<<<<< HEAD
=======
    def Set(model):
        model.PV = pyo.Set() # this will return a set
    
    def Parameters(model, Set, Hours):
        model.pv_eff = pyo.Param(Set) #this returns efficiency to each PV
        model.Sollar_Power = pyo.Param(Hours)  # this returns sower power for each hour in our model
    # def Efficiency(Set):
    #     return pyo.Param(Set) #this returns efficiency to each PV
    
    # def Sollar_Power(Hours):
    #     return pyo.Param(Hours) # this returns sower power for each hour in our model

    def Power(model,t,n):
        model.P_PV = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
        return (model.Sollar_Power[t] * model.pv_eff[n] == model.P_pv[t,n])
    
    def Net(model):
        model.P_pv_net = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
# --------------------------------------- Important searche which baterry it wants to charge
    def Batteries(model):
        model.P_pv_bat = pyo.Var(model.HOURS,model.PV, within = pyo.NonNegativeReals)
# ----------------------------------------

    def Demand(model):
        model.P_pv_demand = pyo.Var(model.HOURS, model.PV,within = pyo.NonNegativeReals)


>>>>>>> b84b25036dd974f68e72fe27e1783896b92019db

    # def create_Sets(self):
    #     self.model.PV = pyo.Set() # model.PV
    #     return {self.type: self.model.PV}
    
    # def create_Param(self):
    #     self.model.pv_eff = pyo.Param(self.model.PV)
    #     self.model.P_solar = pyo.Param(self.model.HOURS) #time series with solar energy
    #     return {'Efficiency':self.model.pv_eff, 'Energy':self.model.P_solar}

    
    # def create_Var(self):
    #     self.model.P_pv = pyo.Var(self.model.HOURS, self.model.PV, within = pyo.NonNegativeReals)
    #     self.model.P_pv_net = pyo.Var(self.model.HOURS, self.model.PV, within = pyo.NonNegativeReals)
    #     self.model.P_pv_bat = pyo.Var(self.model.HOURS,self.model.PV, within = pyo.NonNegativeReals)
    #     self.model.P_pv_demand = pyo.Var(self.model.HOURS, self.model.PV,within = pyo.NonNegativeReals)
    #     return {'Power':self.model.P_pv, 'Power_to_Net':self.model.P_pv_net, 'Power_to_battery':self.model.P_pv_bat, 'Power_to_demand':self.model.P_pv_demand}


    # def solar_rule(self,t,n):
    #     return self.model.P_solar[t] * self.model.pv_eff[n] == self.model.P_pv[t,n]
    # def pv_rule(self,t,n):
    #     return self.model.P_pv[t,n] == self.model.P_pv_bat[t,n] + self.model.P_pv_net[t,n] + self.model.P_pv_demand[t,n]
    
<<<<<<< HEAD
class Bat:
=======
class Batteries:
    pass

class House_Batterie(Batteries):
    def Parameters(model):
        model.starting_SOC = pyo.Param()
        model.E_bat_max = pyo.Param() #capacity of battery
        model.bat_ch_eff = pyo.Param() #charging efficiency of battery
        model.bat_dis_eff = pyo.Param() #discharging efficiency of battery
        model.c_rate_ch = pyo.Param() #maximal charging power (max charging power = c_rate * E_bat_max)
        model.c_rate_dis = pyo.Param() #maximal discharging power (max discharging power = c_rate * E_bat_max)


>>>>>>> b84b25036dd974f68e72fe27e1783896b92019db
    # def __init__(self,model, quantity):
    #     self.model = model
    #     self.quantity = quantity

    # def create_Param(self):
    #     self.model.starting_SOC = pyo.Param()
    #     self.model.E_bat_max = pyo.Param() #capacity of battery
    #     self.model.bat_ch_eff = pyo.Param() #charging efficiency of battery
    #     self.model.bat_dis_eff = pyo.Param() #discharging efficiency of battery
    #     self.model.c_rate_ch = pyo.Param() #maximal charging power (max charging power = c_rate * E_bat_max)
    #     self.model.c_rate_dis = pyo.Param() #maximal discharging power (max discharging power = c_rate * E_bat_max)

    # def create_Var(self):
    #     self.model.SOC = pyo.Var(self.model.HOURS, within = pyo.NonNegativeReals, bounds=(0, 1))
    #     self.model.P_bat_ch = pyo.Var(self.model.HOURS,within = pyo.NonNegativeReals)
    #     self.model.P_bat_dis = pyo.Var(self.model.HOURS, within = pyo.NonNegativeReals)
    #     self.model.P_bat_net = pyo.Var(self.model.HOURS,within = pyo.NonNegativeReals)
    #     self.model.P_bat_demand = pyo.Var(self.model.HOURS, within = pyo.NonNegativeReals)
    #     self.model.K_ch = pyo.Var(self.model.HOURS, domain = pyo.Binary)
    #     self.model.K_dis = pyo.Var(self.model.HOURS, domain = pyo.Binary)

    # def battery_rule(model,t,m):
    #     if t ==1:
    #          return model.SOC[t,m] == model.starting_SOC[m]
    #     else:
    #         return model.SOC[t,m] == model.SOC[t-1,m] + (model.P_bat_ch[t-1] * model.bat_ch_eff[m] -
    #                                                   model.P_bat_dis[t-1]/model.bat_dis_eff[m]) * model.time_step / model.E_bat_max[m]
    # def charge_limit(model,t,m):
    #     return model.P_bat_ch[t,m] <= model.E_bat_max[m] * model.c_rate_ch[m] * model.K_ch[t,m]
    
    # def discharge_limit(model,t,m):
    #     return model.P_bat_dis[t,m] <= model.E_bat_max[m] * model.c_rate_dis[m] * model.K_dis[t,m]
    
    # def keys_rule(model,t,m):
    #     return model.K_ch[t,m] + model.K_dis[t,m] <= 1
    
class General:
    def sell_energy(model,t):
        return model.E_sell[t] == model.P_sell[t] * model.time_step
    
    def buy_energy(model,t):
        return model.E_buy[t] == model.P_net_demand[t] * model.time_step

# endregion
# ---------------------------------------------------------------------------------------------------------------------