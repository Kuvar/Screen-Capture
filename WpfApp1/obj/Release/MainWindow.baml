   M S B A M L   `   `   `  ����   WpfApp1!    ScreenCaptureApp.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -     �Name$  Window�� � Icon$ icon.ico�� � Title$ Capture Screen��$
��Auto��$
��Auto�� � SizeToContent$ WidthAndHeight=�.���� $��10q�$	��Top=� �� Orientation$ Vertical=�.+��� -    ��Name$ 
radioPanel��$	��Top=�$ 
Horizontal=�.+�	�  	� 	GroupName$ Os��$��5q�$��Save As Image�� d� 	IsChecked �� *SaveAsImage+	� $ Os��$��47,5,5,5q�$��Save In Word�� �� *
SaveInWord+�� -   $ SaveInWordPanel��$	��Top=�y��� *SaveInWordPanel+.+�� $��Stretch=� � RowDefinitions �� $
��auto ��� $
��auto �	 � ColumnDefinitions	 �� $
��auto ��� $
��auto �.+�	� -   
 	�Name$
 
chkNewFile��$��0��$��0��$ SaveInWordPanel��$��5,15,0,0q�$��New File�� �� *	NewFile+�� -    ��Name$ txtFileName��$��0��$��1�� �� 	IsEnabled .  � Language$ en-US�$��25��$
��Left=�$	��150��$��	10,10,0,0q�$��3q����� *
FileName+ �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�{��� *	NewFile+	� -   $
 chkExistingFile��$��1��$��0��$ SaveInWordPanel��$��5,15,0,0q�$��Existing File�� �� *ExistingFile+�� -    ��Name$ BrowseButton��$��1��$��1��$��Browse a file��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��150�� J� RenderTransformOrigin$ 1.047,0.821&�$��25��{��� *ExistingFile+���� *BrowseButton_Click+