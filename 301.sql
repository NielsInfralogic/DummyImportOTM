Select ID, Long_Name, Long_Name_Add, Short_Name, Udg_Dato, Pro_PageCount, Pro_SecName, Pro_SecRem, Zones, Pro_Editions_Unique from DATA
Where CAST(UDG_DATO AS Date) = '2023-01-30'
And Short_Name = 'A44'
And Zones = '1'
