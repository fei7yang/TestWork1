D i m   o b j X L  
 S e t   o b j X L   =   W S c r i p t . C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
 o b j X L . D i s p l a y A l e r t s = F A L S E  
 m a g e r P a t h   =   W S c r i p t . A r g u m e n t s ( 0 )  
 S e t   e b   =   o b j X L . W o r k b o o k s . A d d ( m a g e r P a t h )  
 S e t   s h e e t 1   =   e b . W o r k s h e e t s ( 1 )  
 s P a t h   =   W S c r i p t . A r g u m e n t s ( 1 )  
 s a v e P a t h   =   W S c r i p t . A r g u m e n t s ( 2 )  
 s h e e t n n a m e   =   W S c r i p t . A r g u m e n t s ( 3 )  
 c o m p S t r   =   " ~ $ "  
 v b s S t r   =   " . v b s "  
 S e t   o F s o   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )    
 S e t   o F o l d e r   =   o F s o . G e t F o l d e r ( s P a t h )              
 S e t   o F i l e s   =   o F o l d e r . F i l e s    
 F o r   E a c h   o F i l e   I n   o F i l e s  
 	 I f   I n s t r ( 1 , o F i l e . N a m e   , v b s S t r , 1 ) = 0   T h e n  
 	 	 I f   I n s t r ( 1 , m a g e r P a t h   , o F i l e . N a m e , 1 ) = 0   T h e n  
 	 	 	 I f   I n s t r ( 1 , o F i l e . N a m e , c o m p S t r   , 1 ) = 0   T h e n  
 	 	 	 	 ' M s g B o x   o F i l e  
 	 	 	 	 m   =   0  
 	 	 	 	 S e t   w s b   =   o b j X L . W o r k b o o k s . A d d ( o F i l e )  
 	 	 	 	 F o r   m   =   w s b . S h e e t s . C o u n t   T o   1   S t e p   - 1  
 	 	 	 	 i f   w s b . S h e e t s ( m ) . N a m e   =   s h e e t n n a m e   t h e n   w s b . S h e e t s ( m ) . C o p y   e b . W o r k s h e e t s ( 1 )  
 	 	 	 	 N e x t  
 	 	 	 	 w s b . C l o s e  
 	 	 	 E n d   I f  
 	 	 E n d   I f  
 	 E n d   I f  
 N e x t  
 ' M s g B o x   s a v e P a t h  
 e b . S a v e A s   s a v e P a t h  
 e b . C l o s e  
 o b j X L . Q u i t  
 