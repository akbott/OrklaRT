   M S B A M L   `   `   `  ����   SelectionPane    SelectionPane.Selection   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     -     >�Name$  ucrSelection��$��Stretch=�$��Stretch=�$
��Auto��$	��465��.��� .+��� -   $��tbcSelection��.@��� -   $��tbiSelection��$��Utvalg��.��(� $��Stretch=�$	��Top=� (� Stretch$
 None=�$	��450��.��� .+�9� -   $��expVariants��$
��Left=�$��0,0,0,0q�$	��Top=�.���� $
��Left=�$	��Top=�$	��450��.��� -   $��grdVariants�����	#FFE5E5E5$��0,5,0,0q� � ColumnDefinitions �� $
��319* ��� $
��119* �.+��� -   $��lblVariantName��$��Variant Navn��$��25��$	��110��$��10,8,0,0q�$
��Left=�$	��Top=��� -   $��txtVariantName��$��22��$	��200��$��	120,8,0,0q�$
��Left=�$	��Top=�$��2���� -   $��lblDescription��$��Variant Beskrivelse��$��25��$	��110��$��	10,35,0,0q�$
��Left=�$	��Top=��� -	   $��txtDescription��$��22��$	��200��$��
120,35,0,0q�$
��Left=�$	��Top=�$��2���� -
   $��btnOK��$��Lagre��$��22��$��55��$��
150,65,0,0q�$
��Left=�$	��Top=��� -   $��	btnCancel��$��
kansellere��$��22��$��60��$��
210,65,0,0q�$
��Left=�$	��Top=��� -   $��btnRun��$��
kjøre(F8)��$��22��$��51��$��25,0,0,0q�$
��Left=�$	��Top=��� -   $��cboSelectionVariant�� �� SelectedValuePath$ ID�� �� DisplayMemberPath$ VariantName�� �� 
IsEditable .  �� IsTextSearchEnabled . $��22��$	��200��$��80,1,0,0q�$
��Left=�$	��Top=��� -   $��btnSaveSelection��$��Lagre Variant��$��22��$��75��$��290,0,0,486q�$
��Left=�$��Bottom=��� -   $��btnDeleteSelection�� � ToolTip$ Delete��$��22��$��20��$��	375,0,0,0q�$
��Left=�$	��Top=�.���� $��
Delete.jpg��$
��Fill=��� -   $��
lblMessage��$�� ��$��0,25,0,0q�$y�Hidden=��� -   $��btnYes��$	��Yes��$��22��$��50��$��
165,50,0,0q�$
��Left=�$	��Top=�$y�Hidden=��� -   $��btnNew��$	��New��$��22��$��50��$��
220,50,0,0q�$
��Left=�$	��Top=�$y�Hidden=��� -   $��
btnCancel1��$��Canel��$��22��$��50��$��
275,50,0,0q�$
��Left=�$	��Top=�$y�Hidden=��� -   $��grpEasyDateSelection��$��Enkel dato selection��$
��Left=�$��0,80,0,0q�$	��Top=�$	��450��.��� -   $��grdEasyDateSelection�����	#FFE5E5E5.+��� -   $��lblWeek��$
��Uke:��$��25��$��50��$��10,8,0,0q�$
��Left=�$	��Top=��� -   $��txtWeek��$��20��$��25��$��	50,10,0,0q�$
��Left=�$	��Top=��� -   $��lblMonth��$��Måned:��$��25��$��50��$��81,8,0,0q�$
��Left=�$	��Top=��� -   $��txtMonth��$��20��$ MM��$��25��$��
130,10,0,0q�$
��Left=�$	��Top=��� -   $��lblYear��$
��År:��$��25��$��25��$��	158,8,0,0q�$
��Left=�$	��Top=��� -   $��txtYear��$��20��$
 YYYY��$��35��$��
190,10,0,0q�$
��Left=�$	��Top=��� -   $��
lblWarning��$�� ��$��	0,130,0,0q�$y�Visible=��� -   $��grpSelection��$��	Selection��$
��Left=�$��	0,170,0,0q�$	��Top=�$	��450��.��� -   $��grdSelection�����	#FFE5E5E5.+��� -   $��lblName1��$��25��$	��200��$��5,10,0,0q�$
��Left=�$	��Top=��� -    $��lblName2��$��25��$	��200��$��5,35,0,0q�$
��Left=�$	��Top=��� -!   $��lblName3��$��25��$	��200��$��5,60,0,0q�$
��Left=�$	��Top=��� -"   $��lblName4��$��25��$	��200��$��5,85,0,0q�$
��Left=�$	��Top=��� -#   $��lblName5��$��25��$	��200��$��	5,110,0,0q�$
��Left=�$	��Top=��� -$   $��lblName6��$��25��$	��200��$��	5,135,0,0q�$
��Left=�$	��Top=��� -%   $��lblName7��$��25��$	��200��$��	5,160,0,0q�$
��Left=�$	��Top=��� -&   $��lblName8��$��25��$	��200��$��	5,185,0,0q�$
��Left=�$	��Top=��� -'   $��lblName9��$��25��$	��200��$��	5,210,0,0q�$
��Left=�$	��Top=��� -(   $��	lblName10��$��25��$	��200��$��	5,235,0,0q�$
��Left=�$	��Top=��� -)   $!��tbiMultipleSelectionOptions��$��Multippel Utvalg��$y�Hidden=�.��� ���	#FFE5E5E5.+��� -*   $��btnPaste��$��Lim inn��$��22��$��50��$��0,0,0,0q�$
��Left=�$	��Top=��� -+   $��	btnDelete��$��Slett��$��22��$��50��$��55,0,0,0q�$
��Left=�$	��Top=��� -,   $��lblScreenID��$
��Left=�$	��Top=�$y�Hidden=�&   System.Windows.Controls.DataGrid  --   $��dgSelectSingleValues��$��0,25,0,0q�   AutoGenerateColumns .  	   RowHeaderWidth$	 20Q�
   	RowHeight$
 23Q�   CanUserAddRows .    CanUserDeleteRows . $
��Left=�$	��Top=�   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   
Visibility$ Hidden=�1  +System.Windows.Controls.DataGridBoundColumn   Binding ��  �� Path$ ID�