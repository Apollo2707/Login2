   M S B A M L   `   `   `  ����L  HLogin Application, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:Login_ApplicationLogin_Application  5   	   "    Login_Application.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   1localclr-namespace:Login_Application     5   	     � Title$  
MainWindow��5   	   $
��1034��6   $
��2570��6*   .��5      � .+�5   
   ��  �� Orientation$ 
Horizontal=�6   .+�5      �� -   5        $��Login��5      $
��1024��6%   $
��1280��63   $��0q�6A   .��5      � .+�5      �� $��10q�6   .+�5      �� $��45(�6    $/��)Welcome to JBSA Lackland Education Office��6/   $��Center=�6`   �� 5      $��40(�6    $6��0Please fillout this form and take a seat outside��6/   $��Center=�6g   �� 5      $ 
Horizontal=�6!   .+�5      �� -   5        $��cboAppt��5   #   $
��Left=�62   $	��Top=�6N   $��43��6f   .@�5      �� $	��Yes��6+   $��43��69   $
��Left=�6D   �� 5      $��No��6+   $��43��68   $
��Left=�6C   5      �� 5      $��Have an appointment?��6    5       �� 5!      $ 
Horizontal=�6!   .+�5"      �� $��First Name: ��6$   �� 5#      -   5        $��txtFirstName��5#   "   $
��Left=�66   $��23��6Q    �� TextWrapping$
 Wrap=�6]   $	��Top=�6r   $	��120��6�   5$      �� 5%      $ 
Horizontal=�6!   .+�5&      �� $��Last Name: ��6$   �� 5'      -   5        $��txtLastName��5'   #   $
��Left=�66   $��23��6Q   $
 Wrap=�6]   $��1,0,0,0q�6q   $	��Top=�6�   $	��120��6�   5(      �� 5)      $ 
Horizontal=�6!   .+�5*      �� $��Pay Grade:  ��6$   �� 5+      -   5         ��Name$ cboGrade��5+   #   $
��Left=�65   $	��Top=�6Q   $��40��6i   .@�5,      �� $��E1��6+   $��40��69   $��Center=�6E   �� 5-      $��E2��6+   $��40��69   $��Center=�6E   �� 5.      $��E3��6+   $��40��69   $��Center=�6E   �� 5/      $��E4��6+   $��40��69   $��Center=�6E   �� 50      $��E5��6+   $��40��69   $��Center=�6E   �� 51      $��E6��6+   $��40��69   $��Center=�6E   �� 52      $��E7��6+   $��40��69   $��Center=�6E   �� 53      $��E8��6+   $��40��69   $��Center=�6E   �� 54      $��E9��6+   $��40��69   $��Center=�6E   �� 55      $��O1��6+   $��40��69   $��Center=�6E   �� 56      $��O2��6+   $��40��69   $��Center=�6E   �� 57      $��O3��6+   $��40��69   $��Center=�6E   �� 58      $��O4��6+   $��40��69   $��Center=�6E   �� 59      $��O5��6+   $��40��69   $��Center=�6E   �� 5:      $��O6��6+   $��40��69   $��Center=�6E   �� 5;      $��Mr��6+   $��40��69   $��Center=�6E   �� 5<      $��Ms��6+   $��40��69   $��Center=�6E   5=      5>      �� 5?      $ 
Horizontal=�6!   .+�5@      �� $��Branch:       ��6$   �� 5A      -   5        $ 	cboBranch��5A   #   $
��Left=�66   $	��Top=�6R   $��60��6j   $��1,0,0,0q�6u    J� RenderTransformOrigin$ 
0.75,1.273&�6�   .@�5B      �� $
��USAF��6,   $��60��6;   $��Center=�6F   �� 5C      $	��USA��6+   $��60��69   $��Center=�6D   �� 5D      $	��USN��6+   $��60��69   $��Center=�6D   �� 5E      $
��USMC��6+   $��60��6:   $��Center=�6E   �� 5F      $
��USCG��6+   $��60��6:   $��Center=�6E   5G      5H      �� 5I      $ 
Horizontal=�6!   .+�5J      �� $��Status:        ��6$   �� 5K      -   5        $ 	cboStatus��5K   #   $
��Left=�66   $	��Top=�6R   $��2,0,0,0q�6j   $��65��6{   .@�5L      �� $��Active��6+   $��65��6<   $��Center=�6G   �� 5M      $��Guard��6+   $��65��6;   $��Center=�6F   �� 5N      $��Reserve��6+   $��65��6=   $��Center=�6H   �� 5O      $��Retired��6+   $��65��6=   $��Center=�6H   �� 5P      $	��APF��6+   $��65��69   $��Center=�6D   �� 5Q      $	��NAF��6+   $��65��69   $��Center=�6D   �� 5R      $��Family��6+   $��65��6<   $��Center=�6G   5S      5T      �� 5U      $ 
Horizontal=�6!   .+�5V      �� $��Email:         ��6$   �� 5W      -   5         ��Name$ txtEmail��5W   "   $
��Left=�64   $��23��6O   $
 Wrap=�6[   $��2,0,0,0q�6o   $	��Top=�6�   $	��245��6�   5X      �� 5Z      $1��+Please check all the reasons for Your visit��6    �� 5[      $%��(More than one can be selected)��6    �� 5\      -	   5        $��chkInfo��5\      $��General Info��6.   $
��Left=�6E   $	��Top=�6`   �� 5]      -
   5        $��chkCCAF��5]      $
��CCAF��6.   $
��Left=�6=   $	��Top=�6X   �� 5^      -   5        $��chkCommission��5^      $��
Commission��64   $
��Left=�6I   $	��Top=�6d   �� 5_      -   5        $��
chkEdLevel��5_      $��Ed Level Update��61   $
��Left=�6K   $	��Top=�6f   �� 5`      -   5        $��chkInOut��5`      $��In / Out Processing��6/   $
��Left=�6M   $	��Top=�6h   �� 5a      -   5        $��chkPME��5a      $	��PME��6-   $
��Left=�6;   $	��Top=�6V   �� 5b      -   5        $��chkTA��5b      $��TA��6,   $
��Left=�69   $	��Top=�6T   �� 5c      -   5        $��chkVA��5c      $��VA��6,   $
��Left=�69   $	��Top=�6T   �� 5d      -   5        $��	chkAFCOOL��5d      $��AF COOL��60   $
��Left=�6B   $	��Top=�6]   �� 5e      -   5        $��chkWithdraw��5e      $��Withdraw / Reimburse��62   $
��Left=�6Q   $	��Top=�6l   �� 5f      -   5        $��chkOther��5f      $��Other��6/   $
��Left=�6?   $	��Top=�6Z   �� 5g      -   5        $��Submit��5g      $
��Left=�6.   $	��Top=�6J   $��75��6b   6m   5h      5i      5j      �� 5k      -   5        $��Admin��5k      $
��1024��6$   $
��1280��62   $��0q�6?   .��5l      � .+�5m      �� $��10q�6   .+�5n      �� $��Center=�6    $��ADMIN��6=   $��40(�6J   �� 5o      $��Center=�6'   $�� ��6D   $��30(�6L   �� 5p      $ 
Horizontal=�6!   .+�5q      �� $��Admin Intials��6$   �� 5r      -   5        $��	txtIntial��5r   "   $��50��63   $�� ��6>   $��	10,0,50,0q�6F   �� 5s      -   5        $��Save to Excel and close��5s   !   6C   5t      � 5u      $��0q�6   $
��1260��6&   .+�5v      �� $
��Left=�6$   $��0��6?   $��Intial��6O   �� 5w      $
��Left=�6$   $��1��6?   $
��Appt��6P   �� 5x      $
��Left=�6$   $��2��6?   $��
First Name��6P   �� 5y      $
��Left=�6$   $��3��6?   $��	Last Name��6O   �� 5z      $
��Left=�6$   $��4��6@   $��Grade��6Q   �� 5{      $
��Left=�6$   $��5��6?   $��Branch��6O   �� 5|      $
��Left=�6$   $��6��6?   $��Status��6O   �� 5}      $
��Left=�6$   $��7��6?   $��Email��6O   �� 5~      $
��Left=�6$   $��8��6?   $��Reasons for Visit��6O    � ColumnDefinitions 5      �� 5�      $��40 �6/   �� 5�      $��40 �6/   �� 5�      $��80 �6/   �� 5�      $��80 �6/   �� 5�      $��40 �6/   �� 5�      $��50 �6/   �� 5�      $��40 �6/   �� 5�      $	��180 �6/   �� 5�      $��* �6/   5�      �� 5�      -   5        $��lstListPeople��5�      $	��4,4q�63   $��2q�6@   $	��861��6T   ��5�      �� 5�      .R�5�   "   �  5�   &   �� 5�   *   $��40 �6;   �� 5�   *   $��40 �6;   �� 5�   *   $��80 �6;   �� 5�   *   $��80 �6;   �� 5�   *   $��40 �6;   �� 5�   *   $��50 �6;   �� 5�   *   $��40 �6;   �� 5�   *   $	��180 �6;   �� 5�   *   $��* �6;   .+�5�   &   � $��0��6+    5�   *   �� 5�   .   $��40 �6?   .+�5�   *   �� -   5        $
��Left=�5�   1   $��0��6L   $
��Help��6x   y�6]   �� *Vis+6�   �� 5�   *   $
��Left=�65   $��0��6P   y�6`   �� *RVis+��6|   �� *Helped+5�   '   �� 5�   &   $
��Left=�60   $��1��6K   ��6\   �� *Appt+�� 5�   &   $
��Left=�60   $��2��6K   ��6\   �� *	FirstName+�� 5�   &   $
��Left=�60   $��3��6L   ��6\   �� *
LastName+�� 5�   &   $
��Left=�60   $��4��6K   ��6\   �� *Grade+�� 5�   &   $
��Left=�60   $��5��6K   ��6[   �� *Branch+�� 5�   &   $
��Left=�60   $��6��6K   ��6[   �� *Status+�� 5�   &   $
��Left=�60   $��7��6K   ��6[   �� *Email+�� 5�   &   $
��Left=�60   $��8��6K   ��6[   �� *	Reasons+5�   #   5�      5�      5�      5�      5�      5�      5�      5�      