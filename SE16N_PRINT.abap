*&---------------------------------------------------------------------*
*& Report ZCA_SE16N_MAIL
*&---------------------------------------------------------------------*
*& Über dieses Programm werden SE16N Varianten ausgewählt und im
*& Anschluss die Daten entsprechend selektiert und via Email versendet
*&
*& Erstellt: 22.08.2025
*& Entwickler: Guenther Wohlfart (extern)
*&---------------------------------------------------------------------*
REPORT zca_se16n_mail.

INCLUDE zca_se16n_mail_cl.

PARAMETERS: p_vari  TYPE variant OBLIGATORY,
            p_rmail RADIOBUTTON GROUP r1 DEFAULT 'X' USER-COMMAND uc1,
            p_rprin RADIOBUTTON GROUP r1.

SELECTION-SCREEN BEGIN OF BLOCK selscr1 WITH FRAME TITLE TEXT-002.
  PARAMETERS:
    p_mail   TYPE ad_smtpadr MODIF ID m1,
    p_subj   TYPE so_obj_des MODIF ID m1,
    p_body_1 TYPE char255    LOWER CASE MODIF ID m1,
    p_body_2 TYPE char255    LOWER CASE MODIF ID m1,
    p_body_3 TYPE char255    LOWER CASE MODIF ID m1.
SELECTION-SCREEN END OF BLOCK selscr1.

SELECTION-SCREEN BEGIN OF BLOCK selscr2 WITH FRAME TITLE TEXT-003.
  PARAMETERS:
    p_pdest TYPE sypri_pdest MATCHCODE OBJECT prin MODIF ID p1,
    p_pimm  TYPE pri_params-primm AS CHECKBOX MODIF ID p1.
SELECTION-SCREEN END OF BLOCK selscr2.

SELECTION-SCREEN BEGIN OF BLOCK selscr3 WITH FRAME TITLE TEXT-004.
  PARAMETERS p_self_1 TYPE se16n_seltab-field.
  SELECT-OPTIONS  s_sel_1 FOR sy-datum.
  PARAMETERS p_self_2 TYPE se16n_seltab-field.
  SELECT-OPTIONS s_sel_2 FOR sy-datum.
  PARAMETERS p_self_3 TYPE se16n_seltab-field.
  SELECT-OPTIONS s_sel_3 FOR sy-datum.
SELECTION-SCREEN END OF BLOCK selscr3.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_vari.
  DATA lt_return TYPE STANDARD TABLE OF ddshretval.

  SELECT * FROM varit INTO TABLE @DATA(lt_values) WHERE langu = @sy-langu AND
                                                        report = 'SE16N_BATCH'.
  CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'
    EXPORTING
      retfield        = 'VARIANT'
      value_org       = 'S'
    TABLES
      value_tab       = lt_values
      return_tab      = lt_return
    EXCEPTIONS
      parameter_error = 1
      no_values_found = 2
      OTHERS          = 3.

  IF sy-subrc NE 0.
    RETURN.
  ENDIF.

  READ TABLE lt_return INTO DATA(ls_return) INDEX 1.
  IF sy-subrc = 0.
    p_vari = ls_return-fieldval.
  ENDIF.

AT SELECTION-SCREEN.
  IF sy-ucomm = 'ONLI'.
    IF p_rmail = abap_true AND
       ( p_mail   IS INITIAL OR
         p_subj   IS INITIAL OR
         p_body_1 IS INITIAL ).
      MESSAGE TEXT-005 TYPE 'E'.
    ELSEIF p_rprin = abap_true AND
           p_pdest IS INITIAL.
      MESSAGE TEXT-006 TYPE 'E'.
    ENDIF.
  ENDIF.

AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.
    CASE screen-group1.
      WHEN 'M1'.
        IF p_rmail = abap_true.
          screen-active = 1.
          screen-invisible = 0.
        ELSE.
          screen-active = 0.
          screen-invisible = 1.
        ENDIF.
        MODIFY SCREEN.
      WHEN 'P1'.
        IF p_rprin = abap_true.
          screen-active = 1.
          screen-invisible = 0.
        ELSE.
          screen-active = 0.
          screen-invisible = 1.
        ENDIF.
        MODIFY SCREEN.
    ENDCASE.
  ENDLOOP.

START-OF-SELECTION.
  NEW lcl_email_report( )->process( iv_vari   = p_vari
                                    iv_rmail  = p_rmail
                                    iv_rprin  = p_rprin
                                    iv_pdest  = p_pdest
                                    iv_pimm   = p_pimm
                                    iv_mail   = p_mail
                                    iv_subj   = p_subj
                                    iv_body_1 = p_body_1
                                    iv_body_2 = p_body_2
                                    iv_body_3 = p_body_3
                                    iv_self_1 = p_self_1
                                    iv_self_2 = p_self_2
                                    iv_self_3 = p_self_3
                                    it_sel_1  = s_sel_1[]
                                    it_sel_2  = s_sel_2[]
                                    it_sel_3  = s_sel_3[] ).


*&---------------------------------------------------------------------*
*& Include          ZCA_SE16N_MAIL_CL
*&---------------------------------------------------------------------*

CLASS lcl_email_report DEFINITION FINAL.
  PUBLIC SECTION.
    TYPES lty_tt_date TYPE RANGE OF sy-datum.
    TYPES lty_tt_out  TYPE STANDARD TABLE OF se16n_output.

    METHODS process IMPORTING iv_vari   TYPE variant
                              iv_rmail  TYPE xfeld
                              iv_rprin  TYPE xfeld
                              iv_pdest  TYPE sypri_pdest
                              iv_pimm   TYPE pri_params-primm
                              iv_mail   TYPE ad_smtpadr
                              iv_subj   TYPE so_obj_des
                              iv_body_1 TYPE char255
                              iv_body_2 TYPE char255
                              iv_body_3 TYPE char255
                              iv_self_1 TYPE se16n_seltab-field
                              iv_self_2 TYPE se16n_seltab-field
                              iv_self_3 TYPE se16n_seltab-field
                              it_sel_1  TYPE lty_tt_date
                              it_sel_2  TYPE lty_tt_date
                              it_sel_3  TYPE lty_tt_date.

    METHODS create_csv_xstring IMPORTING it_data        TYPE STANDARD TABLE
                                         it_out         TYPE lty_tt_out
                                         iv_tab         TYPE se16n_tab
                                         it_sel_tab     TYPE se16n_or_seltab_t OPTIONAL
                               RETURNING VALUE(rv_data) TYPE xstring.
    METHODS : build_top_of_page IMPORTING iv_tab           TYPE se16n_tab
                                          it_sel_tab       TYPE se16n_or_seltab_t
                                RETURNING VALUE(ro_header) TYPE REF TO cl_salv_form_layout_grid.
    METHODS : create_top_text IMPORTING is_sel_tab TYPE se16n_seltab
                                        iv_field   TYPE scrtext_m
                              EXPORTING ev_text    TYPE char200.
    METHODS : build_seltab_mail IMPORTING it_sel_tab       TYPE se16n_or_seltab_t
                                          iv_field         TYPE se16n_tab
                                RETURNING VALUE(rv_string) TYPE string.

ENDCLASS.


CLASS lcl_email_report IMPLEMENTATION.
  METHOD process.
    TYPES: BEGIN OF lty_indxkey,
             fix  TYPE c LENGTH 5,
             guid TYPE c LENGTH 15,
           END OF lty_indxkey.

    TYPES: BEGIN OF ty_table_ref,
             table_name TYPE rsparams-selname,
             table_ref  TYPE REF TO data,
           END OF ty_table_ref.

    " --- Importing-Parameter same as in SE16N_BATCH ---
    DATA i_tab           TYPE se16n_tab.
    DATA i_no_txt        TYPE c LENGTH 1.
    DATA i_max           LIKE sy-tabix.
    DATA i_line          TYPE c LENGTH 1.
    DATA i_clnt          TYPE c LENGTH 1.
    DATA i_vari          TYPE slis_vari.
    DATA i_guid          TYPE c LENGTH 20.
    DATA i_tech          TYPE c LENGTH 1.
    DATA i_cwid          TYPE c LENGTH 1.
    DATA i_roll          TYPE c LENGTH 1.
    DATA i_conv          TYPE c LENGTH 1.
    DATA i_lget          TYPE c LENGTH 1.
    DATA i_add_f         TYPE c LENGTH 40.
    DATA i_add_on        TYPE c LENGTH 1.
    DATA i_uname         LIKE sy-uname.
    DATA i_hana          TYPE c LENGTH 1.
    DATA i_dbcon         TYPE dbcon_name.
    DATA i_ojkey         TYPE tswappl.
    DATA i_mincnt        LIKE sy-tabix.
    DATA i_fda           TYPE se16n_fda.
    DATA i_formul        TYPE gtb_formula.
    DATA i_exread        TYPE c LENGTH 1.
    DATA i_exwrit        TYPE c LENGTH 1.
    DATA i_exname        TYPE se16n_lt_name.
    DATA i_exunam        TYPE syuname.

    " --- Table parameters
    DATA lt_or_selfields TYPE se16n_or_t.
    DATA lt_out          TYPE STANDARD TABLE OF se16n_output.
    DATA lt_curr_add_up  TYPE STANDARD TABLE OF se16n_output.
    DATA lt_quan_add_up  TYPE STANDARD TABLE OF se16n_output.
    DATA lt_sum_up       TYPE STANDARD TABLE OF se16n_output.
    DATA lt_group_by     TYPE STANDARD TABLE OF se16n_output.
    DATA lt_order_by     TYPE STANDARD TABLE OF se16n_output.
    DATA lt_toplow       TYPE STANDARD TABLE OF se16n_seltab.
    DATA lt_sortorder    TYPE STANDARD TABLE OF se16n_seltab.
    DATA lt_aggregate    TYPE STANDARD TABLE OF se16n_seltab.
    DATA lt_having       TYPE STANDARD TABLE OF se16n_seltab.

    DATA lt_valutab      TYPE TABLE OF rsparams.
    DATA lr_pay_data     TYPE REF TO data.
    DATA lt_csv          TYPE truxs_t_text_data.
    DATA lv_csv          TYPE string.
    DATA lv_csv_x        TYPE xstring.
    DATA lt_csv_x        TYPE solix_tab.
    DATA lr_table        TYPE REF TO data.
    DATA lr_line         TYPE REF TO data.
    DATA lv_lines        TYPE sy-tabix.
    DATA ls_indxkey      TYPE lty_indxkey.
    DATA lt_body         TYPE bcsy_text.
    DATA ls_pctl         TYPE alv_s_pctl.
    DATA lt_tables       TYPE HASHED TABLE OF ty_table_ref WITH UNIQUE KEY table_name.

    FIELD-SYMBOLS <lv_any>       TYPE any.
    FIELD-SYMBOLS <lt_any_table> TYPE ANY TABLE.
    FIELD-SYMBOLS <lt_pay_data>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS <lt_table>     TYPE STANDARD TABLE.
    FIELD-SYMBOLS <ls_line>      TYPE any.

    " --- Read variant ---
    CALL FUNCTION 'RS_VARIANT_CONTENTS'
      EXPORTING
        report  = 'SE16N_BATCH'
        variant = iv_vari
      TABLES
        valutab = lt_valutab
      EXCEPTIONS
        OTHERS  = 1.

    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv3.
      RETURN.
    ENDIF.

    " --- Fill parameters and tabs from variant ---
    LOOP AT lt_valutab INTO DATA(ls_val) WHERE low IS NOT INITIAL AND low <> 'MANDT'.
      CASE ls_val-kind.
        WHEN 'P'. " Parameter
          ASSIGN (ls_val-selname) TO <lv_any>.
          IF sy-subrc = 0.
            <lv_any> = ls_val-low.
          ENDIF.

        WHEN 'S'. " Select Option
          ASSIGN (ls_val-selname) TO <lt_any_table>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          DATA(lo_table) = CAST cl_abap_tabledescr( cl_abap_typedescr=>describe_by_data( <lt_any_table> ) ).
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          DATA(lo_line) = CAST cl_abap_structdescr( lo_table->get_table_line_type( ) ).
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          DATA(lv_name) = lo_line->get_relative_name( ).
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          READ TABLE lt_tables INTO DATA(ls_table) WITH TABLE KEY table_name = ls_val-selname.
          IF sy-subrc = 0.
            lr_table = ls_table-table_ref.
          ELSE.
            CREATE DATA lr_table TYPE TABLE OF (lv_name).
            IF sy-subrc <> 0.
              CONTINUE.
            ENDIF.
            INSERT VALUE #( table_name = ls_val-selname
                            table_ref  = lr_table ) INTO TABLE lt_tables.
          ENDIF.

          ASSIGN lr_table->* TO <lt_table>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          CREATE DATA lr_line TYPE (lv_name).
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          ASSIGN lr_line->* TO <ls_line>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          ASSIGN COMPONENT 'FIELD' OF STRUCTURE <ls_line> TO FIELD-SYMBOL(<lv_field>).
          IF sy-subrc = 0.
            <lv_field> = ls_val-low.
          ENDIF.

          APPEND <ls_line> TO <lt_table>.
          <lt_any_table> = <lt_table>.
      ENDCASE.
    ENDLOOP.

    ls_indxkey-fix  = 'SE16N'.
    ls_indxkey-guid = i_guid.
    IMPORT lt_or_selfields = lt_or_selfields FROM DATABASE indx(al) ID ls_indxkey.

    " If date fields are selected, add them with current date to selection
    READ TABLE lt_or_selfields ASSIGNING FIELD-SYMBOL(<ls_or_selfields>) INDEX 1.
    IF sy-subrc = 0.
      LOOP AT it_sel_1 INTO DATA(ls_sel_1).
        APPEND VALUE se16n_seltab( field  = iv_self_1
                                   sign   = ls_sel_1-sign
                                   option = ls_sel_1-option
                                   low    = ls_sel_1-low
                                   high   = COND #( WHEN ls_sel_1-high = '00000000' THEN '' ELSE ls_sel_1-high ) )
               TO <ls_or_selfields>-seltab.
      ENDLOOP.
      LOOP AT it_sel_2 INTO DATA(ls_sel_2).
        APPEND VALUE se16n_seltab( field  = iv_self_2
                                   sign   = ls_sel_2-sign
                                   option = ls_sel_2-option
                                   low    = ls_sel_2-low
                                   high   = COND #( WHEN ls_sel_2-high = '00000000' THEN '' ELSE ls_sel_2-high ) )
               TO <ls_or_selfields>-seltab.
      ENDLOOP.
      LOOP AT it_sel_3 INTO DATA(ls_sel_3).
        APPEND VALUE se16n_seltab( field  = iv_self_3
                                   sign   = ls_sel_3-sign
                                   option = ls_sel_3-option
                                   low    = ls_sel_3-low
                                   high   = COND #( WHEN ls_sel_3-high = '00000000' THEN '' ELSE ls_sel_3-high ) )
               TO <ls_or_selfields>-seltab.
      ENDLOOP.

      DATA(lt_seltab) =  <ls_or_selfields>-seltab.
    ENDIF.
    " --- Call SE16N_INTERFACE  ---
    CALL FUNCTION 'SE16N_INTERFACE'
      EXPORTING
        i_tab                 = i_tab
        i_no_txt              = i_no_txt
        i_max_lines           = i_max
        i_line_det            = i_line
        i_display             = abap_false
        i_clnt_spez           = i_clnt
        i_variant             = i_vari
        i_tech_names          = i_tech
        i_cwidth_opt_off      = i_cwid
        i_scroll              = i_roll
        i_no_convexit         = i_conv
        i_layout_get          = i_lget
        i_add_field           = i_add_f
        i_add_fields_on       = i_add_on
        i_uname               = i_uname
        i_hana_active         = i_hana
        i_dbcon               = i_dbcon
        i_ojkey               = i_ojkey
        i_formula_name        = i_formul
        i_mincnt              = i_mincnt
        i_fda                 = i_fda
        i_extract_read        = i_exread
        i_extract_write       = i_exwrit
        i_extract_name        = i_exname
        i_extract_uname       = i_exunam
      IMPORTING
        e_line_nr             = lv_lines
        e_dref                = lr_pay_data
      TABLES
        it_or_selfields       = lt_or_selfields
        it_output_fields      = lt_out
        it_add_up_curr_fields = lt_curr_add_up
        it_add_up_quan_fields = lt_quan_add_up
        it_sum_up_fields      = lt_sum_up
        it_having_fields      = lt_having
        it_group_by_fields    = lt_group_by
        it_order_by_fields    = lt_order_by
        it_toplow_fields      = lt_toplow
        it_sortorder_fields   = lt_sortorder
        it_aggregate_fields   = lt_aggregate
      EXCEPTIONS
        no_values             = 1
        OTHERS                = 2.

    IF sy-subrc <> 0 OR lv_lines = 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      RETURN.
    ENDIF.

    ASSIGN lr_pay_data->* TO <lt_pay_data>.
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    " --- Printout
    IF iv_rprin = abap_true.
      TRY.
          cl_salv_table=>factory( EXPORTING list_display = abap_true
                                  IMPORTING r_salv_table = DATA(lo_salv)
                                  CHANGING  t_table      = <lt_pay_data> ).

          lo_salv->get_display_settings( )->set_list_header(
              |System: { sy-sysid } - Tabelle: { i_tab } - Variante: { iv_vari }| ).
          DATA(lo_topofpage) = build_top_of_page( iv_tab         = i_tab
                                                  it_sel_tab  = lt_seltab ).
          lo_salv->set_top_of_list( lo_topofpage ).
          lo_salv->set_top_of_list_print( lo_topofpage ).
          lo_salv->get_functions( )->set_all( ).
          lo_salv->get_columns( )->set_optimize( abap_true ).

          DATA(lo_columns) = lo_salv->get_columns( ).
          LOOP AT lo_columns->get( ) INTO DATA(ls_column).
            READ TABLE lt_out WITH KEY field = ls_column-columnname TRANSPORTING NO FIELDS.
            IF sy-subrc <> 0.
              lo_columns->get_column( ls_column-columnname )->set_technical( abap_true ).
            ENDIF.
          ENDLOOP.

          DATA(lo_print) = lo_salv->get_print( ).
          lo_print->set_print_only( 'N' ).

          CALL FUNCTION 'GET_PRINT_PARAMETERS'
            EXPORTING
              copies                 = '1'
              destination            = iv_pdest
              immediately            = iv_pimm
              no_dialog              = abap_true
            IMPORTING
              out_parameters         = ls_pctl-pri_params
            EXCEPTIONS
              archive_info_not_found = 1
              invalid_print_params   = 2
              invalid_archive_params = 3
              OTHERS                 = 4.
          IF sy-subrc <> 0.
            MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                    WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
          ENDIF.

          lo_print->set_print_control( ls_pctl ).
          lo_print->set_print_parameters_enabled( abap_false ).
          lo_print->set_report_standard_header_on( abap_true  ).
          lo_salv->display( ).

        CATCH cx_salv_not_found INTO DATA(lr_salv_not_found).
          MESSAGE lr_salv_not_found TYPE 'E'.
        CATCH cx_salv_msg INTO DATA(lr_salv_msg).
          MESSAGE lr_salv_msg TYPE 'E'.
      ENDTRY.
    ENDIF.

    " --- Email Versand
    IF iv_rmail = abap_true.

      lv_csv_x = create_csv_xstring( it_data    = <lt_pay_data>
                                     it_out     = lt_out
                                     iv_tab     = i_tab
                                     it_sel_tab = lt_seltab ).

      CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
        EXPORTING
          buffer     = lv_csv_x
        TABLES
          binary_tab = lt_csv_x.

      TRY.
          DATA(lo_send_request) = cl_bcs=>create_persistent( ).

          " Sender (system user)
          DATA(lo_sender) = cl_sapuser_bcs=>create( sy-uname ).
          lo_send_request->set_sender( lo_sender ).

          " Recipient (from selection parameter)
          DATA(lo_recipient) = cl_cam_address_bcs=>create_internet_address( iv_mail ).
          lo_send_request->add_recipient( lo_recipient ).

          " Subject + Body
          APPEND VALUE soli( line = iv_body_1 ) TO lt_body.
          APPEND INITIAL LINE TO lt_body.
          APPEND VALUE soli( line = iv_body_2 ) TO lt_body.
          APPEND INITIAL LINE TO lt_body.
          APPEND VALUE soli( line = iv_body_3 ) TO lt_body.

          DATA(lo_document) = cl_document_bcs=>create_document( i_type    = 'RAW'
                                                                i_subject = iv_subj
                                                                i_text    = lt_body ).

          " Attach CSV
          lo_document->add_attachment( i_attachment_type     = 'CSV'
                                       i_attachment_subject  = |{ i_tab }_{ sy-datum }|
                                       i_att_content_hex     = lt_csv_x
                                       i_attachment_filename = |{ i_tab }_{ sy-datum }.csv| ).

          lo_send_request->set_document( lo_document ).
          lo_send_request->send( i_with_error_screen = abap_true ).
          COMMIT WORK.
          MESSAGE text-001 TYPE 'S'.

        CATCH cx_send_req_bcs
              cx_document_bcs
              cx_address_bcs INTO DATA(lo_exc).
          MESSAGE lo_exc->get_text( ) TYPE 'E'.
          RETURN.
      ENDTRY.
    ENDIF.
  ENDMETHOD.

  METHOD create_csv_xstring.
    DATA lv_line       TYPE string.
    DATA lv_value      TYPE string.
    DATA lv_csv_string TYPE string.
    DATA lt_dd03l      TYPE STANDARD TABLE OF dfies.
    DATA lv_tabname    TYPE ddobjname.

    lv_tabname = iv_tab.
    DATA(lv_seltab_string) = build_seltab_mail( EXPORTING it_sel_tab = it_sel_tab
                                                          iv_field   = iv_tab ).
    CALL FUNCTION 'DDIF_FIELDINFO_GET'
      EXPORTING
        tabname        = lv_tabname
        langu          = sy-langu
        all_types      = 'X'
      TABLES
        dfies_tab      = lt_dd03l
      EXCEPTIONS
        not_found      = 1
        internal_error = 2
        OTHERS         = 3.
    IF sy-subrc <> 0.
      " No handler
    ENDIF.

    LOOP AT it_out INTO DATA(ls_out).
      READ TABLE lt_dd03l INTO DATA(ls_dd03l) WITH KEY fieldname = ls_out-field.
      IF sy-subrc = 0.
        DATA(lv_column_text) = ls_dd03l-scrtext_m.
      ELSE.
        lv_column_text = ls_out-field.
      ENDIF.
      IF lv_csv_string IS INITIAL.
        lv_csv_string = lv_column_text.
      ELSE.
        lv_csv_string = |{ lv_csv_string };{ lv_column_text }|.
      ENDIF.
    ENDLOOP.

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<ls_line>).
      CLEAR lv_line.

      LOOP AT it_out INTO ls_out.
        ASSIGN COMPONENT ls_out-field OF STRUCTURE <ls_line> TO FIELD-SYMBOL(<lv_field>).
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        lv_value = <lv_field>.

        " Escape quote in source data with double quotes
        REPLACE ALL OCCURRENCES OF `"` IN lv_value WITH `""`.

        " Always quote field to avoid delimiter / newline issues
        lv_value = |"{ lv_value }"|.

        " Add values with separator
        IF sy-tabix = 1.
          lv_line = lv_value.
        ELSE.
          lv_line = |{ lv_line };{ lv_value }|.
        ENDIF.
      ENDLOOP.
      " Add line to csv string
      CONCATENATE lv_csv_string lv_line INTO lv_csv_string SEPARATED BY cl_abap_char_utilities=>newline.
    ENDLOOP.
*...// Add the selection tab string to the mail string.
    CONCATENATE lv_seltab_string lv_csv_string INTO lv_csv_string  SEPARATED BY cl_abap_char_utilities=>newline.
    " Convert to UTF-8 and xstring
    TRY.
        rv_data = cl_bcs_convert=>string_to_xstring( iv_string     = lv_csv_string
                                                     iv_convert_cp = abap_true
                                                     iv_codepage   = '4110' " UTF-8
                                                     iv_add_bom    = abap_true ).
      CATCH cx_bcs INTO DATA(lo_exc_bcs).
        MESSAGE lo_exc_bcs->get_text( ) TYPE 'E'.

    ENDTRY.
  ENDMETHOD.
*...// Build the top of page with the variant selection data
  METHOD build_top_of_page.
    DATA lv_tabname    TYPE ddobjname.
    DATA lt_dd03l      TYPE STANDARD TABLE OF dfies.
    DATA lv_row        TYPE i.
    DATA lo_label      TYPE REF TO cl_salv_form_label.

    lv_tabname = iv_tab.
    ro_header = NEW cl_salv_form_layout_grid( ).

*...// Get the DDIC field names of selected Table
    CALL FUNCTION 'DDIF_FIELDINFO_GET'
      EXPORTING
        tabname        = lv_tabname
        langu          = sy-langu
        all_types      = abap_true
      TABLES
        dfies_tab      = lt_dd03l
      EXCEPTIONS
        not_found      = 1
        internal_error = 2
        OTHERS         = 3.
    IF sy-subrc <> 0.
      " No handler
    ENDIF.

    LOOP AT it_sel_tab ASSIGNING FIELD-SYMBOL(<ls_sel_tab>).
      lv_row = sy-tabix .
      READ TABLE lt_dd03l INTO DATA(ls_dd03l) WITH KEY fieldname = <ls_sel_tab>-field.
      IF sy-subrc = 0.
        DATA(lv_column_text) = ls_dd03l-scrtext_m.
        create_top_text( EXPORTING is_sel_tab = <ls_sel_tab>
                           iv_field   = ls_dd03l-scrtext_m
                 IMPORTING  ev_text    = DATA(lv_text) ).
*...// Create the text
        lo_label = ro_header->create_label( row = lv_row column = 1 ).
        lo_label->set_text( lv_text ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.
  METHOD create_top_text.
    CONSTANTS lv_include TYPE char10 VALUE 'Include'.
    CONSTANTS lv_exclude TYPE char10 VALUE 'Exclude'.
    CONSTANTS lv_to      TYPE char2  VALUE '-'.
    CONSTANTS lv_till    TYPE char4  VALUE 'TILL'.

    IF is_sel_tab-sign = 'I'.
      DATA(lv_sign) = lv_include.
    ELSE.
      lv_sign = lv_exclude.
    ENDIF.
    IF is_sel_tab-option IS INITIAL.
      IF is_sel_tab-high IS INITIAL.
        DATA(lv_option) = 'EQ' .
      ELSE.
        lv_option = 'BT' .
      ENDIF.
    ELSE.
      lv_option = is_sel_tab-option .
    ENDIF.
    IF is_sel_tab-low IS NOT INITIAL AND is_sel_tab-high IS NOT INITIAL.
      ev_text = |{ iv_field } [{ lv_sign }, { lv_option }]: { is_sel_tab-low } { lv_to } { is_sel_tab-high } |.
    ELSEIF is_sel_tab-low IS NOT INITIAL AND is_sel_tab-high IS INITIAL.
      ev_text = |{ iv_field } [{ lv_sign }, { lv_option }]: { is_sel_tab-low } |.
    ENDIF.

  ENDMETHOD.

  METHOD build_seltab_mail.
    DATA lv_tabname    TYPE ddobjname.
    DATA lt_dd03l      TYPE STANDARD TABLE OF dfies.
    DATA lv_row        TYPE i.
    DATA lo_label      TYPE REF TO cl_salv_form_label.

    lv_tabname = iv_field.

*...// Get the DDIC field names of selected Table
    CALL FUNCTION 'DDIF_FIELDINFO_GET'
      EXPORTING
        tabname        = lv_tabname
        langu          = sy-langu
        all_types      = abap_true
      TABLES
        dfies_tab      = lt_dd03l
      EXCEPTIONS
        not_found      = 1
        internal_error = 2
        OTHERS         = 3.
    IF sy-subrc <> 0.
      " No handler
    ENDIF.

    LOOP AT it_sel_tab ASSIGNING FIELD-SYMBOL(<ls_sel_tab>).
      lv_row = sy-tabix .
      READ TABLE lt_dd03l INTO DATA(ls_dd03l) WITH KEY fieldname = <ls_sel_tab>-field.
      IF sy-subrc = 0.
        DATA(lv_column_text) = ls_dd03l-scrtext_m.
        create_top_text( EXPORTING is_sel_tab = <ls_sel_tab>
                                   iv_field   = ls_dd03l-scrtext_m
                         IMPORTING ev_text    = DATA(lv_text) ).
*...// Create the text
        IF rv_string IS INITIAL.
          rv_string = lv_text.
          CONCATENATE rv_string cl_abap_char_utilities=>newline INTO rv_string.
        ELSE.
          CONCATENATE rv_string lv_text cl_abap_char_utilities=>newline INTO rv_string.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.
ENDCLASS.
