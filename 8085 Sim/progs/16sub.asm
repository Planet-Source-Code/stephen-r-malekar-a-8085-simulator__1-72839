org 4000h
	lxi h,l1
	mov a,m
	inx h
	inx h
	sub m
	mov e,a
	lxi h,h1
	mov a,m
	inx h
	inx h
	sbb m
	mov d,a
	mvi a,00h
	jnc store
	mov a,e
	cma
	mov e,a
	mov a,d
	cma
	mov d,a
	inx d
	xra a
	inr a
store:	lxi h,5000h
	mov m,e
	inx h
	mov m,d
	inx h
	mov m,a
	hlt
org 4500h
	l1 db 04h
	h1 db 02h
	l2 db 09h
	h2 db 09h
