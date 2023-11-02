import time

variable = 1

start_time = time.time()

if variable == 1:
    end_time_test =time.time()
    time.sleep(2)
else:
    pass
end_time_other = time.time()


elapsed_time = end_time_other -(end_time_other-end_time_test) - start_time
print(f"Elapsed time: {elapsed_time} seconds")




