org 4000h
	lxi h, 4500h
	mov a,m
	mov b,a
	mvi c,08h
	mvi d,00h
	ora a
loop:	rar
	mov b,a
	mov a,d
	ral
	mov d,a
	mov a,b
	dcr c
	jnz loop
	mov a,d
	sta 5000h
	hlt
org 4500h
	db 63h 	