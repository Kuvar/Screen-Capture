   M S B A M L   `   `   `  ����c  _System.Windows.Interactivity, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35qPclr-namespace:System.Windows.Interactivity;assembly=System.Windows.InteractivitySystem.Windows.Interactivity  5   	    WpfApp1!   ScreenCaptureApp.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ^iPclr-namespace:System.Windows.Interactivity;assembly=System.Windows.Interactivity     5   	   -   5          �Name$  Window��5   	    � Icon$ icon.ico��5   	    � Title$ Screen Capture��5	   	   $
��Auto��5
   	   $
��Auto��5   	    � SizeToContent$ WidthAndHeight=�5   	   .   (System.Windows.Interactivity.Interaction   Triggers 5      /   )System.Windows.Interactivity.EventTrigger  5   
      	EventName$ Closing��6   .   (System.Windows.Interactivity.TriggerBase   Actions. 5      6   0System.Windows.Interactivity.InvokeCommandAction     Command 6$   �� *WindowClosing+   CommandParameter 6F   �� 5      .��5      � .+�5   
   �� $��0q�6   
������ $��1q�66   	 �� CornerRadius$	 0��6J   ��5      �� 5      
 �� EndPoint$
 	0.504,1.5&�6&    �� 
StartPoint$ 
0.504,0.03&�6;   .��5      �  � Color$ #3494e6��6#    � Offset$ 0Q�63   � 5      $ #dbe6f6��6#   $ 0.567Q�63   5      ._�5      ]� .+�5      �� $��10q�6   $	��Top=�6)    �� Orientation$ Vertical=�6A   .+�5      �� -   5         ��Name$ 
radioPanel��5   !   $	��Top=�65   $ 
Horizontal=�6M   .+�5      	�  	� 	GroupName$ Os��6&   $��5q�65   $��Save As Image��6@    d� 	IsChecked 6X   �� *SaveAsImage+	� 5      $ Os��6&   $��47,5,5,5q�65   $��Save In Word��6G    6^   �� *
SaveInWord+5       �� 5!      -   5        $ SaveInWordPanel��5!   !   $	��Top=�6:   y�6R   �� *SaveInWordPanel+.+�5"      � $��Stretch=�6    � RowDefinitions 5#      �� 5$   "   $
��auto �60   �� 5%   "   $
��auto �60    � ColumnDefinitions 5'      �� 5(   "   $
��auto �63   �� 5)   "   $
��auto �63   .+�5+      	� -   5         	�Name$ 
chkNewFile��5+   *   $��0��5,      $��0��5-      $ SaveInWordPanel��5.      $��5,15,0,0q�5/      $��New File��50       51      �� *	NewFile+�� 53      -   5         ��Name$ txtFileName��53   &   $��0��54      $��1��55       �� 	IsEnabled .  � Language$ en-US�58      $��25��59      $
��Left=�5:      $	��150��5;      $��	10,10,0,0q�5<      $��3q�5=      ��56      �� *
FileName+ �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�{�5>      �� *	NewFile+	� 5@      -   5        $ chkExistingFile��5@   *   $��1��5A      $��0��5B      $ SaveInWordPanel��5C      $��5,15,0,0q�5D      $��Existing File��5E       5F      �� *ExistingFile+�� 5H      -   5         ��Name$ BrowseButton��5H   %   $��1��5I      $��1��5J      $��Browse a file��5L      $
��Left=�5M      $��	10,10,0,0q�5N      $	��Top=�5O      $	��150��5P       J� RenderTransformOrigin$ 1.047,0.821&�5Q      $��25��5R      {�5K      �� *ExistingFile+��5S      �� *BrowseButton_Click+5T      5U      5V      5W      5X      5Y      5[      