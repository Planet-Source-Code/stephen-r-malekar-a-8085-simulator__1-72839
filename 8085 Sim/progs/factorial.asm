;Compute factorial of an 8 bit number
;org 4000h
	lxi sp,5500h
	lxi h,b1	
	mov c,m
	mov a,c
	dcr a
	jz end
loop1:	push psw
	call prod
	mov c,l
	pop psw
	dcr a
	jnz loop1
end:	shld 5000h
	hlt
prod:	mvi h,00h
	mov l,h
	mvi d,8h
loop:	dad h
        ral
	jnc nxt
	dad b
nxt:	dcr d
	jnz loop
	ret

org 4500h
 b1  db 6h	