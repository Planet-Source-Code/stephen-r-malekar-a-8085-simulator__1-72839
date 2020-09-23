;Find square root of a given 8 bit number
org 4000h
	lxi h,tb
	mov a,m
	mov d,a
	mvi c,00h
	mvi b,01h
	mov e,b
loop:	cmp b
	jm  acc
	inr e
	inr e
	mov a,e
	add b
	mov b,a
	mov a,d		
	inr c
	jmp loop
acc:	mov a,c
	sta 5000h
	hlt
org 4500h
	tb db 1Ah