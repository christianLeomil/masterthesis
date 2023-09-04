class OBJECTIVE:
    def __init__(self,name_of_instance):
        self.name_of_instance = name_of_instance

        self.list_var = []
        self.list_text_var = []

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input

        #defining energy type to build connections with other componets correctly
        self.super_class = 'objective'