# zcl_file_Data_convert
Abap file data converter

Convert Internal data to File data on the fly.

# Usage Example
### ITAB to XLSX.
```ABAP
 data(lt_xstring) = NEW zcl_filedata_convert( )->itab_2_xlsx( itab ).
```
### XLSX to ITAB.
```ABAP
  NEW zcl_filedata_convert( )->xlsx_2_itab( EXPORTING it_solix = lt_itab" Solix data 
                                            IMPORTING et_data  = lt_target_itab 
                                          ).
```
### ITAB to CSV
```ABAP
data(lt_csv_Data) = NEW zcl_filedata_convert( )->itab_2_csv( itab ).
```

