;Program to check whether a number is odd or even        
        org 4000H                                     
        lxi sp, 5020h                                
        lda start                                     
        call check                                    
        sta start                                     
        hlt                                           
        org 3000h                                     
        start db 05h                                  
        org 7000h                                     
check:  rar                                                
        jnc even                                      
        mvi a, 0ddh                                   
        jmp  return                                   
even:   mvi a, 0eeh                                       
return: ret                                              
