org 4000h
	lxi h, b1
	mov b,m
	mvi c,08h
	inx h	
	mov a,m
	jm lab0
loop1:	ral
	ora a
	jp loop1
lab0:	mov d,a
	mov a,b
	mov b,d
loop:	ora a
	jp lab1	
	mov d,a
	xra b
	mov e,a
	mov a,d
	mov d,e
lab1:	ral
	dcr c
	jnz loop:
	hlt
	b1 db 9h
	b2 db 4h