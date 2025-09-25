#Note that this stopwatch is tailored to my specific needs (rounding, 6 iteration sections, speed_factor),
#   so will need some tweaking when used for other purposes
import json
from time import perf_counter_ns #returns time in nanoseconds

speed_factor = 0.5 #set to 1 for normal functionality

output = {}
for section in range(4,6):
    output[section] = []

    print(f"\nBegin section {section}")

    for _ in range(6):
        input("hit Enter to start")
        start_time = perf_counter_ns()
        input("hit Enter to stop")
        stop_time = perf_counter_ns()

        the_time = round((stop_time - start_time) / 1000000000 * speed_factor, 3)
        print(f"time = {the_time} s")
        output[section].append(the_time)

    with open("savaData.json", "w") as myFile:
        json.dump(output, myFile)

print("end program!")