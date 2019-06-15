# -*- coding: utf-8 -*-
import numpy as np
import matplotlib.pyplot as plt


x = np.linspace(10,20,100)
y1=np.power(x,4)+np.power(x,2)
y2=np.power(2,x)
#x = [0.1,0.2,0.3,0.5,0.8,1,1.5,2,3]
#y1=[43,19,10,5,0.3,0.2,0.15,0.1,0.05]
#y2 = [138,105,91,57,35,24,18,10,8]
plt.figure() 
#plt.plot(x,y1,label="1k transactions",color="blue",marker="^",linewidth=2)
#plt.plot(x,y2,label="10k transactions",color="blue",marker="*",linewidth=2)
plt.plot(x,y1,label="$x^4$",color="red",linewidth=2)
plt.plot(x,y2,label="$2^x$",color="blue",linewidth=2)
#plt.plot(x, y, color="r", linestyle="--", marker="*", linewidth=1.0)
my_x_ticks = np.arange(10, 20, 1)
#my_y_ticks = np.arange(0, 70, 10)
plt.xticks(my_x_ticks)
plt.xlabel("Supprot threshold(%)")
plt.ylabel("Runtime(sec.)")
plt.title("Runtime with threshold")
plt.legend()

#plt.savefig(r'C:\Users\liyun\Desktop\毕设\数据\programming\Runtime with threshold.svg')
plt.show()

y3=np.power(17,4)+np.power(17,2)-np.power(2,17)
print(y3)
