org 4000h
	xra a
	mov d,a
	mvi a,99h
	mov c,a
	lxi h,tb
	mov b,m
	inx h
	sub m
	inr a
	add b
	daa
	jc lab1
	mov b,a
	mov a,c
	sub b
	inr a
	inr d
lab1:	lxi h,5000h
	mov m,a
	inx h
	mov m,d
	hlt
org 4500h
	tb db 46h
	   db 64h
