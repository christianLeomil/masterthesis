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

class Pv(Generator):
    def __init__(self,type_,id,eff,E_in,op_cost,inv_cost,emission):
        super().__init__('pv',id,eff,E_in,op_cost,inv_cost,emission)

    def generation_rule(model,t):
        return model.P_pv[t] == model.P_solar[t] + model.pv_eff
     
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
# region subclasses

# class PV(Generator):
#      def __init__(self,eff,E_in,op_cost,inv_cost,emission):
#           super().__init__(eff,E_in,op_cost,inv_cost,emission)
          
# pv = PV(0.1, 1000, 20, 1000, 100)

# E_out = pv.generation()
# print(E_out)
# print(pv.operation_cost(E_out))
# print(pv.investment())

# endregion