;Program to perform 16 bit division
org 4000h
      lxi h,5000h
      lxi sp,4500h
      mov a,m
      mov b,a
      mvi a,02h
      mvi d,00h
      inx h
      inx h
      mov e,m
lab3: mvi c,08h
      push psw
      inx h
lab1: push h
      mov h,d
      mov l,e
      dad h
      mov d,h
      mov e,l
      pop h
      mov a,d
      sub b
      jc lab2
      inr e
      mov d,a 	
lab2: dcr c
      jnz lab1
      inx h	
      mov m,e
      dcx h  
      dcx h
      dcx h
      mov e,m	
      pop psw
      dcr a
      jnz lab3
      mov m,d
      hlt
org 5000h
      db 05h;
      dw FFFFh;