 
  
 D i m   d e b u g ,   d e v T e s t  
 d e b u g   =   F a l s e  
 d e v T e s t   =   F a l s e  
  
  
 I f   N o t   d e b u g   T h e n  
   O n   E r r o r   R e s u m e   N e x t  
 E n d   I f  
  
 	 I f   W S c r i p t . A r g u m e n t s . C o u n t   <   2   T h e n  
 	 	 M s g B o x   " U s a g e :   S u b s t M a c r o s - E x c e l   < x l s   f i l e >   < f i l e   p a t h > "  
 	 	 W S c r i p t . Q u i t  
 	 E n d   I f  
  
 	 D i m   x l s F i l e N a m e , f i l e P a t h , p r e f i x  
 	 x l s F i l e N a m e   =   W S c r i p t . A r g u m e n t s ( 0 )  
 	 f i l e P a t h   =   W S c r i p t . A r g u m e n t s ( 1 )  
 	 p r e f i x   =   W S c r i p t . A r g u m e n t s ( 2 )  
 	 a l t e r c o d e   =   W S c r i p t . A r g u m e n t s ( 3 )  
 	 r e p o r t t y p e   =   W S c r i p t . A r g u m e n t s ( 4 )   ' A B ( �]z\ONz�^A B h�) G G ( �{t�]z�VA h�) C S ( �N�]
q��
q�c�Speh�)   Z C ( �vPg�m��[��nUS)    
  
  
 	 D i m   f s ,   t x t F i l e ,   l i n e ,   d a t a  
 	 S e t   f s   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
 	 I f   f s . F i l e E x i s t s ( x l s F i l e N a m e )   =   F a l s e   T h e n  
 	     W S c r i p t . E c h o   " F i l e   "   &   x l s F i l e N a m e   &   "   d o e s n ' t   e x i s t ! "  
 	     W S c r i p t . Q u i t  
 	 E n d   I f  
 	  
          
 	 D i m   e x c e l A p p  
 	 S e t   e x c e l A p p   =   C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
 	 e x c e l A p p . V i s i b l e   =   f a l s e  
 	 ' e x c e l A p p . V i s i b l e   =   t r u e  
 	 ' e x c e l A p p . S c r e e n U p d a t i n g   =   d e b u g  
  
 	 D i m   w b  
 	 S e t   w b   =   e x c e l A p p . W o r k b o o k s . O p e n ( x l s F i l e N a m e )  
  
       D i m   s h t , n e w S h e e t , m a x d t e  
       n o w d a t e = n o w    
       F o r   E a c h   s h t   I n   w b . S h e e t s  
       	 	 f l a g   =   F a l s e  
 	 	  
 	 	 i f   r e p o r t t y p e   =   " G G - A "   t h e n  
 	 	     x 1   =   s h t . c e l l s ( 2 8 , 4 )  
 	 	     x 2   =   s h t . c e l l s ( 2 9 , 4 )  
 	 	     x 3   =   s h t . c e l l s ( 3 0 , 4 )  
 	 	     x 4   =   s h t . c e l l s ( 3 1 , 4 )  
 	 	     x 5   =   s h t . c e l l s ( 3 2 , 4 )  
 	 	     x 6   =   s h t . c e l l s ( 2 8 , 1 0 )  
                     x 7   =   s h t . c e l l s ( 2 9 , 1 0 )  
                     x 8   =   s h t . c e l l s ( 3 0 , 1 0 )  
                     x 9   =   s h t . c e l l s ( 3 1 , 1 0 )  
                     x 1 0   =   s h t . c e l l s ( 3 2 , 1 0 ) 	  
 	 	     ' W S c r i p t . E c h o   " G G "  
 	 	 E l s e I f   r e p o r t t y p e   =   " C S "   t h e n  
 	 	     x 1   =   s h t . c e l l s ( 5 2 , 1 2 )  
 	 	     x 2   =   s h t . c e l l s ( 5 3 , 1 2 )  
 	 	     x 3   =   s h t . c e l l s ( 5 4 , 1 2 )  
 	 	     x 4   =   s h t . c e l l s ( 5 2 , 4 4 )  
 	 	     x 5   =   s h t . c e l l s ( 5 3 , 4 4 )  
 	 	     x 6   =   s h t . c e l l s ( 5 4 , 4 4 )  
                     x 7   =   s h t . c e l l s ( 5 2 , 8 2 )  
                     x 8   =   s h t . c e l l s ( 5 3 , 8 2 )  
                     x 9   =   s h t . c e l l s ( 5 4 , 8 2 )  
                     x 1 0   =   " " 	  
 	 	     ' W S c r i p t . E c h o   " C S "  
 	 	 E l s e I f   r e p o r t t y p e   =   " Z C "   t h e n  
 	 	     x 1   =   s h t . c e l l s ( 2 5 , 4 )  
 	 	     x 2   =   s h t . c e l l s ( 2 6 , 4 )  
 	 	     x 3   =   s h t . c e l l s ( 2 7 , 4 )  
 	 	     x 4   =   s h t . c e l l s ( 2 5 , 1 3 )  
 	 	     x 5   =   s h t . c e l l s ( 2 6 , 1 3 )  
 	 	     x 6   =   s h t . c e l l s ( 2 7 , 1 3 )  
                     x 7   =   " "  
                     x 8   =   " "  
                     x 9   =   " "  
                     x 1 0   =   " " 	  
 	 	     ' W S c r i p t . E c h o   " Z C "  
 	 	 e l s e  
 	 	     x 1   =   s h t . c e l l s ( 5 0 , 1 2 )  
 	 	     x 2   =   s h t . c e l l s ( 5 1 , 1 2 )  
 	 	     x 3   =   s h t . c e l l s ( 5 2 , 1 2 )  
 	 	     x 4   =   s h t . c e l l s ( 5 0 , 4 4 )  
 	 	     x 5   =   s h t . c e l l s ( 5 1 , 4 4 )  
 	 	     x 6   =   s h t . c e l l s ( 5 2 , 4 4 )  
                     x 7   =   " "  
                     x 8   =   " "  
                     x 9   =   " "  
                     x 1 0   =   " " 	  
 	 	     ' W S c r i p t . E c h o   " A B "  
 	 	 e n d   i f  
 	 	  
 	 	 ' m a x d t e = d a t e a d d ( " d " , - 3 0 , n o w d a t e )  
 	 	 i f   x 1   < >   " "   t h e n  
 	 	       ' a 1   =   S p l i t ( x 1 , " . " )  
 	 	       ' d t e 1   =   D a t e S e r i a l ( a 1 ( 0 ) , a 1 ( 1 ) , a 1 ( 2 ) )  
 	 	       ' d i f f 1   =   D a t e D i f f ( " d " , d t e 1 , m a x d t e )  
 	 	       i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 1 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	  
 	 	 e n d   i f    
 	 	 i f   x 2   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 2 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 3   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 3 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 4   < >   " "   t h e n  
 	 	       i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 4 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	  
 	 	 e n d   i f    
 	 	 i f   x 5   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 5 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 6   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 6 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 7   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 7 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 8   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 8 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 9   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 9 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 i f   x 1 0   < >   " "   t h e n  
 	 	         i f   T r i m ( a l t e r c o d e )   =   T r i m ( x 1 0 )   t h e n  
 	 	             f l a g   =   T r u e  
 	 	       e n d   i f   	 	 	  
 	 	 e n d   i f    
 	 	 	 	 	 	 	 	  
 	 	 I f   f l a g   T h e n  
 	 	 s h t . C o p y 	  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . S a v e A s   f i l e P a t h   &   " \ "   &   p r e f i x   &   s h t . N a m e  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . S a v e = T r u e  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . C l o s e 	 	  
 	 	 E n d   I f  
 	 N e x t  
 	 ' w b . S a v e = T r u e  
 	 w b . C l o s e   0  
 	 e x c e l A p p . Q u i t   0  
  
 	 D i m   f s o  
 	 S e t   f s o   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
 	 f s o . M o v e F o l d e r   f i l e P a t h ,   f i l e P a t h   & " s u c c e s s "  
  
 	 W S c r i p t . Q u i t  
  
 