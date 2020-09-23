;Sort Array of 8 bit integers
org 4000h
loop1:	lxi h, barray
	mov c,m
	dcr c
	mov m,c
	jm end
	inx h
	mvi e,00h
loop2:	mov a,m
	inx h
	mov b,m
	cmp b
	jp  lab1
	mov d,b
	mov b,a
	mov a,d
	dcx h
	mov m,a
	inx h
	mov m,b
	mvi e,01h
lab1:   dcr c
	jnz loop2:
	mov a,e
	ora a	
	jnz loop1
end:	hlt
org 4500h	
barray  db 0ah
	db 25h
	db 15h
	db 30h
	db 20h
	db 75h
	db 63h
	db 67h
	db 57h
	db 93h
	db 39h
