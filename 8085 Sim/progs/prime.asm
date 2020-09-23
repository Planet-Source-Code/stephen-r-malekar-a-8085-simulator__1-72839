org 4000h
	lxi h,b1
	mov a,m
	cpi 02h
	jm comp
	mov b,a
	ora a
	rar 
	mov c,a
	mov a,b
	mov b,c
	mvi d,00h
loop1:	push psw
 	mov a,b
	cpi 01h	
	jz end
	pop psw
	push psw
	push b
	push d
	call rem	
	mov a,l
	ora a
	jz comp				
	pop d
	pop b
	pop psw
	dcr b
	jmp loop1
comp:	mvi d,01h
end:	lxi h,5000h
	mov m,d
	hlt	
rem:	mvi c,08h
	mvi h,00h
	mov l,h
loop:	dad h
	ral
	jnc lab1
	inx h
lab1:	mov d,a
	mov a,l
	cmp b
	jm lab2:
	sub b
lab2: 	mov l,a
	mov a,d
	dcr c
	jnz loop:
	ret
org 4500h
	b1 db dh