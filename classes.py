class Generator:
    def __init__(self,eff,E_in,op_cost,inv_cost,emission):
        self.eff = eff
        self.E_in= E_in
        self.inv_cost = inv_cost
        self.op_cost = op_cost
        self.emission = emission

    def generation(self):
        E_out = self.E_in * self.eff
        return E_out
    
    def operation_cost(self,E_out):
            cost_total = E_out * self.op_cost
            return cost_total
    
    def investment(self):
         return self.inv_cost

class Consumer:
     def __init__(self,E_demand):
          self.E_demand = E_demand

     def consume(self):
          return self.E_demand
     
class Transformer:
     def __init__(self,eff_in,eff_out,E_in,E_out,E_storage):
          self.eff_in = eff_in
          self.eff_out = eff_out
          self.E_in = E_in
          self.E_out = E_out
          self.E_storage = E_storage

     def charge(self):
          self.E_storage += self.E_in * self.eff_in
          return self.E_storage
     
     def discharge(self):
          self.E_storage -= self.E_out * self.eff_out
          return self.E_storage
     
# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

class PV(Generator):
     def __init__(self,eff,E_in,op_cost,inv_cost,emission):
          super().__init__(eff,E_in,op_cost,inv_cost,emission)
          
pv = PV(0.1, 1000, 20, 1000, 100)

E_out = pv.generation()
print(E_out)
print(pv.operation_cost(E_out))
print(pv.investment())