org 4000h
	lxi h,l1
	mov a,m
	inx h
	inx h
	add m
	sta 5000h
	lxi h,h1
	mov a,m
	inx h
	inx h
	adc m
	sta 5001h
	ral
	ani 1h
	sta 5002h
	hlt
org 4500h
	l1 db 00h
	h1 db 05h
	l2 db 00h
	h2 db 03h
