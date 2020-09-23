;Check whether a given 8 bit number is a palindrome or not
org 4000h
	lxi h,tb
	mov a,m
	mov b,a
	mvi c,cnt
	mvi d,01h
	inx h
loop:   ana m
	jz  cont
	cmp m
	jz  cont
        mvi d,00h
cont:	mov a,b
        inx h
 	dcr c
	jnz loop
	mov a,d
	sta 5000h
	hlt
org 4500h
        tb db a5h
msks	db 81h
	db 42h
	db 24h
	db 18h     
cnt     equ 4h