   M S B A M L   `   `   `  ����G  CPresentation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:Presentation.ViewPresentation.View  5      '    !Presentation.View.WelcomePageView   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      1localclr-namespace:Presentation.View     5      Y UPrism.Unity.Wpf, Version=7.2.0.1422, Culture=neutral, PublicKeyToken=40ee6c3a2184dc59S OPrism.Wpf, Version=7.2.0.1422, Culture=neutral, PublicKeyToken=40ee6c3a2184dc59&prismhttp://prismlibrary.com/   5      !  Prism.Mvvm.ViewModelLocator    AutoWireViewModel  . .��5      �  � RowDefinitions 5   
   �� 5      �� 5      �� 5      �� 5      .+�5   
   �� -   5        $��
textBlock1��5   P   $:��4You successfully have configured Prism into your APP��6   6P   �� 5   
   $��1��63   ��6   �� *ImGoodByeText+63   �� 5   
   -   5        $��FileNameTextBox��5   E   $
��Left=�5      $��43��6-   $��	30,10,0,0q�69    �� TextWrapping$
 Wrap=�6L   $��TextBox��5      $	��Top=�6!   $	��436��69   $��2��6\   �� 5   
   -   5         ��Name$ BrowseButton��5      $��Browse a file��6'   $
��Left=�6?   $��
485,13,0,0q�5      $	��Top=�6%   $	��121��6=    J� RenderTransformOrigin$ 1.047,0.821&�5      $��40��65   $��2��6A   �� 5   
   -   5        $ ShowData��5      $��ShowData��6#   $
��Left=�66   $��
611,13,0,0q�5      $	��Top=�6%   $	��121��6=   $ 1.047,0.821&�5      $��40��65   $��2��6A   �� 5   
   $��3��6   .+�5      &   System.Windows.Controls.DataGrid  -   5          Name$ dtgExcel��5   $   $	��200��6   68   5       �� 5!   
   -   5        $��Dialog��5!      $
��Left=�6"   $��
485,58,0,0q�6=   $��2��6Q   $	��Top=�6^   $	��121��6v   $��31��6�   �� 5"   
   -   5        $��ReadExcelMapper��5"      $
��Left=�6+   $��
611,58,0,0q�6F   $��2��6Z   $	��Top=�6g   $	��121��6   $��31��6�   6�   5#      5$      