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
       D i m   s h t , n e w S h e e t  
       F o r   E a c h   s h t   I n   w b . S h e e t s  
 	 	 ' s h t . C o p y  
 	 	 s h t . C o p y  
 	 	 ' A c t i v e W o r k b o o k . S a v e A s   f i l e P a t h   &   " \ "   &   s h t . N a m e  
 	 	 ' s h t . S a v e A s   f i l e P a t h   &   " \ "   &   s h t . N a m e  
 	 	 ' w b . S a v e A s   x l s F i l e N a m e   &   " . "   &   f s . G e t E x t e n s i o n N a m e ( x l s F i l e N a m e )  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . S a v e A s   f i l e P a t h   &   " \ "   &   p r e f i x   &   s h t . N a m e  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . S a v e = T r u e  
 	 	 e x c e l A p p . A c t i v e W o r k b o o k . C l o s e  
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