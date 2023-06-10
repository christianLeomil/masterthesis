#region superclasses

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