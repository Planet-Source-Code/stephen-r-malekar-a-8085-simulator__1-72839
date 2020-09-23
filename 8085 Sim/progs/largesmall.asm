;Find the largest and smallest number
org 4000h
	lxi h,barray
	mov c,m
	inx h
	mov d,m
	mov e,d
loop:	mov a,m
	cmp d
	jm lab1
	mov d,a
lab1:	cmp e
	jp lab2
	mov e,a
lab2:	inx h
	dcr c
	jnz loop
	lxi h,5000h
	mov m,d
	inx h
	mov m,e
	hlt
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
       