org 4000h
	lxi h, b1
	mov a,m
	inx h
	mov b,m
	mvi c,08h
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
	hlt
org 4500h
	b1 db 9h
	b2 db 4h

	
