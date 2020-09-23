;Find square of a given 8 bit number
org 4000h
	lxi h,tb
	mov c,m
	mvi a,00h
	mvi b,01h
loop:	dcr c
	jm lab1
	add b
	inr b
	inr b
	jmp loop
lab1:	sta 5000h
	hlt
org 4500h
	tb db 9h