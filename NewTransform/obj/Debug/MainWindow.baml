   M S B A M L   `   `   `  ����G  CNewTransform, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:NewTransformNewTransform  5   	       NewTransform.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ,localclr-namespace:NewTransform     5   	   J FWpfAnimatedGif, Version=1.4.18.0, Culture=neutral, PublicKeyToken=null,gif"http://wpfanimatedgif.codeplex.com  5   	   -   5          � Title$  Transform 2.0��5	   	   $	��335��6   $	��335��6/   $	��335��6?   $	��450��6L   $	��450��6[   $	��450��6j    � Icon$ Images/right-arrow.png��6�   .��5
      � .+�5   
   �� .@�5      �� ��5      �� 5       �� Orientation$ 
Horizontal=�6!   .+�5      �� $��Images/envelope.png��6    �� 5      $��3q�6$   	Letters68   5      ��5      �� 5      ���	#FF468B9B.+�5      �� -   5         ��Name$ xlsxDropPanel��5   %   $
��Left=�6<   $	��109��6W   $	��Top=�6d   $	��218��6|   
������� J� 	AllowDrop . $��0,0,0,0q�6�   .+�5      �� $��25��6$   $��25��60   $��94,25,99,20q�6;   $��Images/excel.png��6P   �� 5      -   5         ��Name$ 	xlsxLabel��5   $   $��*.xlsx��67   $��Center=�6H   $��87,-20,90,20q�6e   $��25��6{   $��41��6�   �� 5      -   5        $ xlsxLabelTitle��5   $   $�� ��6<   $��Center=�6G   $��
0,-11,0,20q�6d   $��25��6x   $	��218��6�   {�.  6�   5      �� 5      -   5        $ docDropPanel��5   %   $
��Left=�6;   $	��109��6V   $��224,-109,0,0q�6c   $	��Top=�6y   $	��218��6�   
������� . .+�5      �� $
��Left=�6$   $��25��6?   $��25��6K   $��	95,25,0,0q�6V   $	��Top=�6i   $��Images/word.png��6�   �� 5      -   5        $ docLabel��5   $   $��*.doc(docx)��66   $��
70,0,77,20q�6L   �� 5      -   5        $ docLabelTitle��5   $   $��25��6;   $�� ��6G   $��
0,-12,0,20q�6R   {�.  6f   5      �� 5      -   5        $ SaveFolderPanel��5   %   $
��Left=�6>   $��65��6Y   $��0,6,0,0q�6e   $	��Top=�6v   $	��442��6�   
������� . .+�5       �� $��25��6$   $��25��60   $��Images/folder.png��6;   $
��Fill=�6V   $��18,17,399,0q�6e   �� 5!      -	   5        $ SaveFolderLabel��5!   =   $��Save directory��6$   $��
43,-22,7,0q�6V   5"      �� 5#      -
   5         ��Name$ imgGenerate��5#       $
��Left=�65   "  WpfAnimatedGif.ImageBehavior   RepeatBehavior$ Forever��6P      AnimatedSource$' !Images/Double Ring-1.5s-200px.gif��6{   $��40��6�   $��
379,47,0,0q�6�   $	��Top=�6�   $��55��6�   $'��!Images/Double Ring-1.5s-200px.gif��6  $
��Fill=�6.  {�.  	 K� ZIndex$	 -1��6O  �� 5$      -   5        
 ��Name$
 generatedButton��5$   4   $��Generate��6!   $
��Left=�6M   $��	0,-82,0,0q�6h   $	��Top=�6{   $	��436��6�   $��85��6�   
���������5%      
� 5&   "   .��5'   &   �  � Color$ 	#FF1019DC��63    � Offset$ 1Q�6E   � 5(   &   $ 	#FF2A86B2��63   5)   #   5+      5,      5.      �� 5/      ��50      �� 51      $ 
Horizontal=�6!   $��80��6:   .+�52      A� $��10��6"   $��10��6.   
���   �69   �� 53      $��Images/right-arrow.png��6    $��12��6@   �� 54      $��10��6$   $��10��60   
���   �6;   �� 55      $��3q�6$   	Convert68   56      ��58      �� 59      -   5        $ ConvertPanel��59   !    . ���	#FF468B9B.+�5:      �� -   5        $ 
imgConvert��5:       $ Forever��64   $ Images/giphy.gif��6_   $
��Left=�6�   $	��270��6�   $	��436��6�   $��Images/giphy.gif��6�   $
��Fill=�6�   $��0,0,0,0q�6�   $y�Hidden=�6  � 5;      -   5         �Name$ ConvertProgressBar��5;   &   $��20��6B   $��0, -19, 0, 0q�6N   ���	#FF5E4674���	#FF468B9B$y�Hidden=�6�   ��6{   Z� 6�   �� 5<      -   5        $ ConvertLabelPanel��5<   %   $��0, -300, 0, 0q�6@   $	��170��6W   $��25��6c   .+�5=      �� -   5        $ ConvertLabelText��5=   $   $'��!Drop or click for convert file(s)��6>   
�������6j   5>      5?      5A      5B      5C      5D      