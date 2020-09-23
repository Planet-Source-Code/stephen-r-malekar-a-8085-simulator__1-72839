;Find product of two 8 bit numbers
org 4000h
	lxi h,tb
	mov c,m
	inx h
	mov a,m
	mvi h,00h
	mov l,h
	mvi d,8h
loop:	ral
	dad h
	jnc nxt
	dad b
nxt:	dcr d
	jnz loop
	shld 5000h
	hlt
org 4500h
tb 	db 45h
	db 47h	