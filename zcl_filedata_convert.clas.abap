CLASS ZCL_FILEDATA_CONVERT DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    METHODS:
      itab_2_xlsx            IMPORTING VALUE(it_data)     TYPE STANDARD TABLE
                             RETURNING VALUE(rv_xstring)  TYPE xstring,
      itab_2_csv             IMPORTING it_data            TYPE STANDARD TABLE
                             RETURNING VALUE(rt_data)     TYPE string_table,
      itab_2_tsv             IMPORTING it_data            TYPE STANDARD TABLE
                             RETURNING VALUE(rt_data)     TYPE string_table,
      itab_2_separated_file  IMPORTING it_data            TYPE STANDARD TABLE
                                       iv_separator       TYPE char1
                             RETURNING VALUE(rt_data)     TYPE string_table,
      struc_2_separated_file IMPORTING is_data            TYPE any
                                       iv_separator       TYPE char1
                             RETURNING VALUE(rs_data)     TYPE string,
      struc_2_csv            IMPORTING is_data            TYPE any
                             RETURNING VALUE(rs_data)     TYPE string,
      struc_2_tsv            IMPORTING is_data            TYPE any
                             RETURNING VALUE(rs_data)     TYPE string,
      separated_file_2_struc IMPORTING is_data            TYPE string
                                       iv_separator       TYPE char1
                             EXPORTING VALUE(es_data)     TYPE any,
      csv_2_struc            IMPORTING is_data            TYPE string
                             EXPORTING VALUE(es_data)     TYPE any,
      tsv_2_struc            IMPORTING is_data            TYPE string
                             EXPORTING VALUE(es_data)     TYPE any,
      separated_file_2_itab  IMPORTING it_data            TYPE string_table
                                       iv_separator       TYPE char1
                             EXPORTING VALUE(et_data)     TYPE STANDARD TABLE,
      csv_2_itab             IMPORTING it_data            TYPE string_table
                             EXPORTING VALUE(et_data)     TYPE STANDARD TABLE,
      tsv_2_itab             IMPORTING it_data            TYPE string_table
                             EXPORTING VALUE(et_data)     TYPE STANDARD TABLE,
      xlsx_2_itab            IMPORTING it_solix           TYPE solix_tab
                             EXPORTING VALUE(et_data)     TYPE STANDARD TABLE.
  PROTECTED SECTION.
  PRIVATE SECTION.
    METHODS:
      convert_date_2_internal IMPORTING iv_ext_date          TYPE char10
                              RETURNING VALUE(rv_int_date)   TYPE sydatum,
      convert_time_2_internal IMPORTING VALUE(iv_ext_time)   TYPE char8
                              RETURNING VALUE(rv_int_data)   TYPE syuzeit,
      struc_to_struc          IMPORTING value(iv_struc_from) TYPE any
                              EXPORTING VALUE(ev_struc_to)   TYPE any.

ENDCLASS.



CLASS ZCL_FILEDATA_CONVERT IMPLEMENTATION.


  METHOD convert_date_2_internal.

    CASE iv_ext_date+2(1).
      WHEN '.'."DD
        rv_int_date = iv_ext_date+6(4) && iv_ext_date+3(2) && iv_ext_date(2).
      WHEN '/'"MM
        OR '-'.
        rv_int_date = iv_ext_date+6(4) && iv_ext_date(2) && iv_ext_date+3(2).
      WHEN OTHERS.
        IF iv_ext_date+4(1) = '.' " 'YYYY?'
        OR iv_ext_date+4(1) = '/'
        OR iv_ext_date+4(1) = '-'.
          rv_int_date = iv_ext_date(4) && iv_ext_date+5(2) && iv_ext_date+8(2).
        ELSE.

          RETURN.

        ENDIF.
    ENDCASE.

* Check the valid date.
    CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
      EXPORTING
        date   = rv_int_date
      EXCEPTIONS
        OTHERS = 2.

    IF syst-subrc <> 0.
      rv_int_date = iv_ext_date.
    ENDIF.
  ENDMETHOD.


  METHOD convert_time_2_internal.

    IF iv_ext_time+6(2) >= 60.

      iv_ext_time+7(1) = iv_ext_time+6(1).
      iv_ext_time+6(1) = 0.

    ENDIF.

    TRY.
        cl_abap_timefm=>conv_time_ext_to_int( EXPORTING time_ext               = iv_ext_time    " External Represenation of Time
                                                        is_24_allowed          = 'X'   " Is 24:00 permitted?
                                              IMPORTING time_int               = rv_int_data    " Internal Represenation of Time
                                            ).

      CATCH cx_abap_timefm_invalid.    "

        rv_int_data = iv_ext_time.

    ENDTRY.

  ENDMETHOD.


  METHOD csv_2_itab .

    CHECK et_data IS SUPPLIED.

    me->separated_file_2_itab( EXPORTING it_data      = it_data
                                         iv_separator = ','
                               IMPORTING et_data      = et_data
                             ).

  ENDMETHOD.


  METHOD csv_2_struc .
    CHECK es_Data is SUPPLIED.
    me->separated_file_2_struc( EXPORTING is_data      = is_data
                                          iv_separator = ','
                                IMPORTING es_data      = es_data
                              ).


  ENDMETHOD.


  METHOD itab_2_csv .

    rt_data = me->itab_2_separated_file( it_data      = it_data
                                         iv_separator = ','
                                       ).

  ENDMETHOD.


  METHOD itab_2_separated_file .

    CHECK iv_separator IS NOT INITIAL.

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<lfs_data>).

      APPEND me->struc_2_separated_file( is_data      = <lfs_data>
                                         iv_separator = iv_separator
                                       )
          TO rt_data.
    ENDLOOP.

  ENDMETHOD.


  METHOD itab_2_tsv .

    rt_data = me->itab_2_separated_file( it_data      = it_data
                                         iv_separator = cl_abap_char_utilities=>horizontal_tab

                                       ).

  ENDMETHOD.


  METHOD itab_2_xlsx .


    TRY.

        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_table)
                                 CHANGING t_table      = it_data
                              ).

        DATA(lt_fcat) = cl_salv_controller_metadata=>get_lvc_fieldcatalog( r_columns      = lo_table->get_columns( )
                                                                           r_aggregations = lo_table->get_aggregations( )
                                                                         ).


        DATA(lo_result) = cl_salv_ex_util=>factory_result_data_table( r_data         = REF #( it_data )
                                                                      t_fieldcatalog = lt_fcat
                                                                    ).

        cl_salv_bs_tt_util=>if_salv_bs_tt_util~transform( EXPORTING xml_type      = if_salv_bs_xml=>c_type_xlsx
                                                                    xml_version   = cl_salv_bs_a_xml_base=>get_version( )
                                                                    r_result_data = lo_result
                                                                    xml_flavour   = if_salv_bs_c_tt=>c_tt_xml_flavour_export
                                                                    gui_type      = if_salv_bs_xml=>c_gui_type_gui
                                                          IMPORTING xml           = rv_xstring
                                                        ).


      CATCH cx_salv_msg.

    ENDTRY.

  ENDMETHOD.


  METHOD separated_file_2_itab .

    CHECK et_data IS SUPPLIED.

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<lfs_data>).

      APPEND INITIAL LINE TO et_data ASSIGNING FIELD-SYMBOL(<lfs_data_line>).

      me->separated_file_2_struc( EXPORTING is_data      = <lfs_data>
                                            iv_separator = iv_separator
                                  IMPORTING es_data      = <lfs_data_line>
                                ).


    ENDLOOP.

  ENDMETHOD.


  METHOD separated_file_2_struc .

    CHECK es_data IS SUPPLIED.


    SPLIT is_data
       AT iv_separator
     INTO TABLE DATA(lt_data_fields).

    LOOP AT lt_data_fields INTO DATA(ls_data_field).
      ASSIGN COMPONENT sy-tabix OF STRUCTURE es_data TO FIELD-SYMBOL(<lfs_field>).

      IF syst-subrc <> 0
      OR <lfs_field> IS NOT ASSIGNED.
        EXIT.
      ENDIF.

      <lfs_field> = ls_data_field.
    ENDLOOP.

  ENDMETHOD.


  METHOD struc_2_csv .

    rs_data = me->struc_2_separated_file( is_data      = is_data
                                          iv_separator = ','
                                        ).

  ENDMETHOD.


  METHOD struc_2_separated_file .
    DATA:
      lv_field_string TYPE string.

    DO.
      ASSIGN COMPONENT syst-index OF STRUCTURE is_data TO FIELD-SYMBOL(<lfs_field>).
      IF syst-subrc <> 0
      OR <lfs_field> IS NOT ASSIGNED.
        EXIT.
      ENDIF.

      lv_field_string = <lfs_field>.

      IF rs_data IS INITIAL.
        rs_data = lv_field_string.
      ELSE.
        rs_data = rs_data && iv_separator && lv_field_string.
      ENDIF.
    ENDDO.

  ENDMETHOD.


  METHOD struc_2_tsv .

    rs_data = me->struc_2_separated_file( is_data      = is_data
                                          iv_separator = cl_abap_char_utilities=>horizontal_tab
                                        ).
  ENDMETHOD.


  METHOD tsv_2_itab .

    CHECK et_data IS SUPPLIED.
    me->separated_file_2_itab( EXPORTING it_data      = it_data
                                         iv_separator = cl_abap_char_utilities=>horizontal_tab
                               IMPORTING et_data      = et_data
                             ).

  ENDMETHOD.


  METHOD tsv_2_struc .

    me->separated_file_2_struc( EXPORTING is_data      = is_data
                                          iv_separator = cl_abap_char_utilities=>horizontal_tab
                                IMPORTING es_data      = es_data
                              ).
  ENDMETHOD.


  METHOD xlsx_2_itab .

    FIELD-SYMBOLS:
      <lfs_tab>    TYPE STANDARD TABLE.
    CHECK et_data IS SUPPLIED.

    TRY .
        DATA(lo_excel_ref) = NEW cl_fdt_xl_spreadsheet( document_name = 'FILE.XLSX'
                                                        xdocument     = cl_bcs_convert=>solix_to_xstring( it_solix ) ) .

      CATCH cx_fdt_excel_core.
        RETURN.
    ENDTRY .

    lo_excel_ref->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = DATA(worksheet_names) ).

    DATA(lo_data_ref) = lo_excel_ref->if_fdt_doc_spreadsheet~get_itab_from_worksheet( worksheet_names[ 1 ] ).

    ASSIGN lo_data_ref->* TO <lfs_tab>.

    LOOP AT <lfs_tab> ASSIGNING FIELD-SYMBOL(<lfs_data>) FROM 2.
*
      APPEND INITIAL LINE TO et_data ASSIGNING FIELD-SYMBOL(<lfs_data_line>).
*
      me->struc_to_struc( EXPORTING iv_struc_from = <lfs_data>
                          IMPORTING ev_struc_to   = <lfs_data_line>
                        ).

    ENDLOOP.

  ENDMETHOD.
  METHOD struc_to_struc.

    DO.

      ASSIGN COMPONENT syst-index OF STRUCTURE iv_struc_from TO FIELD-SYMBOL(<lfs_field>).
      IF syst-subrc <> 0
      OR <lfs_field> IS NOT ASSIGNED.
        EXIT.
      ENDIF.
      CHECK <lfs_field> IS NOT INITIAL.
      ASSIGN COMPONENT syst-index OF STRUCTURE ev_struc_to TO FIELD-SYMBOL(<lfs_field2>).
      IF syst-subrc <> 0
      OR <lfs_field2> IS NOT ASSIGNED.
        EXIT.
      ENDIF.

      CASE CAST cl_abap_datadescr( cl_abap_typedescr=>describe_by_data( <lfs_field2> ) )->get_data_type_kind( <lfs_field2> ).

        WHEN cl_abap_typedescr=>typekind_date.

          <lfs_field2> = me->convert_date_2_internal( CONV #( <lfs_field> ) ).

          CONTINUE.

        WHEN   cl_abap_typedescr=>typekind_time.

          CONDENSE <lfs_field>.

          <lfs_field2> = me->convert_time_2_internal( CONV #( <lfs_field> ) ).

          CONTINUE.
      ENDCASE.
      <lfs_field2> = <lfs_field>.

    ENDDO.

  ENDMETHOD.

ENDCLASS.
