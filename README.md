# Malevannyi
ТЕСТИРОВАНИЕ: 
![image](https://user-images.githubusercontent.com/102906241/205466577-c2ecbd68-82c0-46a7-bd34-a070df325565.png)
![image](https://user-images.githubusercontent.com/102906241/205466618-d5d813e6-426a-4103-bdcb-e1244fe1a810.png)

ПРОФИЛИРОВАНИЕ:

Время выполнения прежнего метода обработки даты занимает 0.029 секунд
![img_1.png](img_1.png)

Теперь заменим прежний метод на метод, который использует библиотеку Arrow. Результат по времени выполнения получился еще больше - 0.096 секунд
![img_2.png](img_2.png)

Теперь снова заменим метод, но теперь на метод с использованием
библиотеки Maya. Результат лучше, но все равно долго.
![img_3.png](img_3.png)

В конце концов, попробуем применить метод, которой просто берёт срез
из строки с датой. Время его выполнения существенно быстрее всех предыдущих - 0 секунд.
Оставим его в программе, а остальные методы закомментируем
![img_4.png](img_4.png)
