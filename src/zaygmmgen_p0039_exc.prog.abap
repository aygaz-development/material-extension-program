*&---------------------------------------------------------------------*
*& Include zaygmmgen_p0039_exc
*&---------------------------------------------------------------------*
CLASS zbc_cx_excel DEFINITION
  INHERITING FROM cx_static_check
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_t100_message .
    INTERFACES if_t100_dyn_msg .

    CONSTANTS:
      BEGIN OF zbc_cx_excel,
        msgid TYPE symsgid VALUE 'ZBC_EXCEL',
        msgno TYPE symsgno VALUE '000',
        attr1 TYPE scx_attrname VALUE '',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF zbc_cx_excel .
    CONSTANTS mc_message_class TYPE msgid VALUE 'ZBC_EXCEL' ##NO_TEXT.
    CONSTANTS:
      BEGIN OF mc_message_types,
        information TYPE msgty VALUE 'I',
        error       TYPE msgty VALUE 'E',
        success     TYPE msgty VALUE 'S',
        warning     TYPE msgty VALUE 'W',
        abort       TYPE msgty VALUE 'A',
        error_all   TYPE char3 VALUE 'AEX',
      END OF mc_message_types .
    CONSTANTS:
      BEGIN OF mc_messages,
        generic                 TYPE msgno VALUE '000', "& & & &
        data_ref_is_a_not_table TYPE msgno VALUE '001', "Data reference is not a table!
        file_ext_must_be_xlsx   TYPE msgno VALUE '002', "File extension must be XLSX!
      END OF mc_messages .
    CONSTANTS mc_no_data_ref_is_a_not_table TYPE msgno VALUE '' ##NO_TEXT.

    METHODS constructor
      IMPORTING
        !textid   LIKE if_t100_message=>t100key OPTIONAL
        !previous LIKE previous OPTIONAL
        !msgv1    TYPE symsgv1 OPTIONAL
        !msgv2    TYPE symsgv2 OPTIONAL
        !msgv3    TYPE symsgv3 OPTIONAL
        !msgv4    TYPE symsgv4 OPTIONAL .
    METHODS get_message_as_bapiret2
      RETURNING
        VALUE(rs_bapiret2) TYPE bapiret2 .

    METHODS if_message~get_text
        REDEFINITION .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zbc_cx_excel IMPLEMENTATION.


  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    CALL METHOD super->constructor
      EXPORTING
        previous = previous.
    CLEAR me->textid.
    IF textid IS INITIAL.
      if_t100_message~t100key = zbc_cx_excel .
    ELSE.
      if_t100_message~t100key = textid.
    ENDIF.
  ENDMETHOD.


  METHOD if_message~get_text.
    CALL METHOD super->if_message~get_text
      RECEIVING
        result = result.
    IF previous IS NOT INITIAL.
      result = result && '(' && previous->get_text( ) && ')'.
    ENDIF.
  ENDMETHOD.


  METHOD get_message_as_bapiret2.
    IF me->previous IS INITIAL.

      rs_bapiret2-id     = me->if_t100_message~t100key-msgid.
      rs_bapiret2-number = me->if_t100_message~t100key-msgno.
      rs_bapiret2-type   = me->if_t100_dyn_msg~msgty.

      IF rs_bapiret2-type IS INITIAL.
        rs_bapiret2-type = me->mc_message_types-error.
      ENDIF.

      rs_bapiret2-message_v1 = me->if_t100_dyn_msg~msgv1.
      rs_bapiret2-message_v2 = me->if_t100_dyn_msg~msgv2.
      rs_bapiret2-message_v3 = me->if_t100_dyn_msg~msgv3.
      rs_bapiret2-message_v4 = me->if_t100_dyn_msg~msgv4.

      MESSAGE ID rs_bapiret2-id
            TYPE rs_bapiret2-type
          NUMBER rs_bapiret2-number
            WITH rs_bapiret2-message_v1
                 rs_bapiret2-message_v2
                 rs_bapiret2-message_v3
                 rs_bapiret2-message_v4
            INTO rs_bapiret2-message.

    ELSE.

      rs_bapiret2-message = me->get_text( ).
      rs_bapiret2-id      = me->mc_message_class.
      rs_bapiret2-number  = me->mc_messages-generic.
      rs_bapiret2-type    = me->if_t100_dyn_msg~msgty.

      IF rs_bapiret2-type IS INITIAL.
        rs_bapiret2-type = me->mc_message_types-error.
      ENDIF.

      rs_bapiret2-message_v1 = rs_bapiret2-message(50).
      rs_bapiret2-message_v2 = rs_bapiret2-message+50(50).
      rs_bapiret2-message_v3 = rs_bapiret2-message+100(50).
      rs_bapiret2-message_v4 = rs_bapiret2-message+150(50).

    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS zbc_cl_excel DEFINITION
  FINAL
  CREATE PRIVATE .

  PUBLIC SECTION.

    CLASS-METHODS create_with_internal_table
      IMPORTING
        !ir_data          TYPE REF TO data
        !it_field_catalog TYPE lvc_t_fcat
        !is_layout        TYPE lvc_s_layo OPTIONAL
        !it_sort          TYPE lvc_t_sort OPTIONAL
        !it_filter        TYPE lvc_t_filt OPTIONAL
        !it_hyperlinks    TYPE lvc_t_hype OPTIONAL
      RETURNING
        VALUE(ro_myself)  TYPE REF TO zbc_cl_excel
      RAISING
        zbc_cx_excel .
    CLASS-METHODS create_with_xml
      IMPORTING
        !iv_xml          TYPE xstring
      RETURNING
        VALUE(ro_myself) TYPE REF TO zbc_cl_excel
      RAISING
        zbc_cx_excel .
    CLASS-METHODS upload_xlsx
      IMPORTING
        !iv_filename     TYPE string OPTIONAL
      EXPORTING
        !ev_filepath     TYPE string
      RETURNING
        VALUE(ro_myself) TYPE REF TO zbc_cl_excel
      RAISING
        zbc_cx_excel .
    METHODS download_xlsx
      IMPORTING
        !iv_filename TYPE string OPTIONAL
      RAISING
        zbc_cx_excel .
    METHODS get_data
      RETURNING
        VALUE(rr_data) TYPE REF TO data
      RAISING
        zbc_cx_excel .
    METHODS get_data_by_worksheet_index
      IMPORTING
        !iv_index           TYPE i
        !iv_data_has_header TYPE abap_bool OPTIONAL
      CHANGING
        VALUE(cr_data)      TYPE REF TO data
      RAISING
        zbc_cx_excel .
    METHODS get_data_by_worksheet_name
      IMPORTING
        !iv_worksheet_name TYPE string
      CHANGING
        VALUE(cr_data)     TYPE REF TO data
      RAISING
        zbc_cx_excel .
    METHODS get_xml
      RETURNING
        VALUE(rv_xml) TYPE xstring
      RAISING
        zbc_cx_excel .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA mr_data TYPE REF TO data .
    DATA mv_xml TYPE xstring .

    CLASS-METHODS check_reference_is_table
      IMPORTING
        !ir_data       TYPE REF TO data
      RETURNING
        VALUE(rv_bool) TYPE abap_bool .
    CLASS-METHODS file_open_dialog
      RETURNING
        VALUE(rv_filepath) TYPE string .
    CLASS-METHODS file_save_dialog
      RETURNING
        VALUE(rv_filepath) TYPE string .
    METHODS map_fields
      IMPORTING
        !ir_xlsx_data TYPE REF TO data
      CHANGING
        !cr_itab_data TYPE REF TO data .
    CLASS-METHODS check_filename
      IMPORTING
        !iv_filename   TYPE string
      RETURNING
        VALUE(rv_bool) TYPE abap_bool .
ENDCLASS.



CLASS zbc_cl_excel IMPLEMENTATION.


  METHOD create_with_internal_table.

    IF abap_true NE check_reference_is_table( ir_data = ir_data ).
      "Data reference is not a table!
      RAISE EXCEPTION TYPE zbc_cx_excel
                MESSAGE ID zbc_cx_excel=>mc_message_class
                      TYPE zbc_cx_excel=>mc_message_types-error
                    NUMBER zbc_cx_excel=>mc_messages-data_ref_is_a_not_table.
    ENDIF.

    ro_myself          = NEW #( ).
    ro_myself->mr_data = ir_data.

    TRY .

        cl_salv_bs_tt_util=>if_salv_bs_tt_util~transform(
          EXPORTING
            xml_type      = if_salv_bs_xml=>c_type_xlsx
            xml_version   = if_salv_bs_xml=>version_26
            xml_flavour   = if_salv_bs_c_tt=>c_tt_xml_flavour_export
            gui_type      = if_salv_bs_xml=>c_gui_type_gui
            r_result_data = cl_salv_ex_util=>factory_result_data_table( r_data         = ir_data
                                                                        t_fieldcatalog = it_field_catalog
                                                                        s_layout       = is_layout
                                                                        t_sort         = it_sort
                                                                        t_filter       = it_filter
                                                                        t_hyperlinks   = it_hyperlinks )
          IMPORTING
            xml           = ro_myself->mv_xml ).

      CATCH cx_root INTO DATA(lx_root).
        "& & & &
        RAISE EXCEPTION TYPE zbc_cx_excel
                  MESSAGE ID zbc_cx_excel=>mc_message_class
                        TYPE zbc_cx_excel=>mc_message_types-error
                      NUMBER zbc_cx_excel=>mc_messages-generic
                        WITH lx_root->get_text( ).
    ENDTRY.

  ENDMETHOD.


  METHOD create_with_xml.
    ro_myself = NEW #( ).
    ro_myself->mv_xml = iv_xml.
  ENDMETHOD.


  METHOD download_xlsx.

    DATA: lv_filepath TYPE string,
          lv_length   TYPE i,
          lt_content  TYPE STANDARD TABLE OF tdline.

    IF iv_filename IS INITIAL.

      lv_filepath = file_save_dialog( ).

      IF lv_filepath IS INITIAL.
        RETURN.
      ENDIF.

    ELSE.
      lv_filepath = iv_filename.
    ENDIF.

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = mv_xml
      IMPORTING
        output_length = lv_length
      TABLES
        binary_tab    = lt_content.

    CALL FUNCTION 'GUI_DOWNLOAD'
      EXPORTING
        bin_filesize            = lv_length
        filename                = lv_filepath
        filetype                = 'BIN'
      TABLES
        data_tab                = lt_content
      EXCEPTIONS
        file_write_error        = 1
        no_batch                = 2
        gui_refuse_filetransfer = 3
        invalid_type            = 4
        no_authority            = 5
        unknown_error           = 6
        header_not_allowed      = 7
        separator_not_allowed   = 8
        filesize_not_allowed    = 9
        header_too_long         = 10
        dp_error_create         = 11
        dp_error_send           = 12
        dp_error_write          = 13
        unknown_dp_error        = 14
        access_denied           = 15
        dp_out_of_memory        = 16
        disk_full               = 17
        dp_timeout              = 18
        file_not_found          = 19
        dataprovider_exception  = 20
        control_flush_error     = 21
        OTHERS                  = 22.

  ENDMETHOD.


  METHOD file_open_dialog.

    DATA: lt_filetable TYPE filetable,
          lv_rc        TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        file_filter             = '(*.xlsx)|*.xlsx'
      CHANGING
        file_table              = lt_filetable
        rc                      = lv_rc
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        OTHERS                  = 5 ).

    IF sy-subrc IS NOT INITIAL.

    ENDIF.

    rv_filepath = VALUE #( lt_filetable[ 1 ]-filename OPTIONAL ).

  ENDMETHOD.


  METHOD file_save_dialog.

    DATA: lv_filename TYPE string,
          lv_path     TYPE string,
          lv_fullpath TYPE string.

    cl_gui_frontend_services=>file_save_dialog(
        EXPORTING
          default_file_name         = 'Export'
          default_extension         = 'xlsx'
          file_filter               = '(*.xlsx)|*.xlsx'
        CHANGING
          filename                  = lv_filename
          path                      = lv_path
          fullpath                  = lv_fullpath
        EXCEPTIONS
          cntl_error                = 1
          error_no_gui              = 2
          not_supported_by_gui      = 3
          invalid_default_file_name = 4
          OTHERS                    = 5
      ).

    rv_filepath = lv_fullpath.

  ENDMETHOD.


  METHOD get_data.
    rr_data = mr_data.
  ENDMETHOD.


  METHOD get_data_by_worksheet_index.

    DATA: lr_xl   TYPE REF TO cl_fdt_xl_spreadsheet,
          lr_doc  TYPE REF TO if_fdt_doc_spreadsheet,

          lr_data TYPE REF TO data.

    FIELD-SYMBOLS: <lt_data>  TYPE table,
                   <lv_value> TYPE any.
    TRY.
        CREATE OBJECT lr_xl
          EXPORTING
            document_name = ''
            xdocument     = mv_xml.

        lr_doc ?= lr_xl.

        lr_doc->get_worksheet_names(
          IMPORTING
            worksheet_names = DATA(lt_worksheets) ).

        DATA(lv_worksheet_name) = VALUE #( lt_worksheets[ iv_index ] OPTIONAL ).

        lr_data = lr_doc->get_itab_from_worksheet( lv_worksheet_name ).

      CATCH cx_fdt_excel_core INTO DATA(lref_excel_core).
        "& & & &
        RAISE EXCEPTION TYPE zbc_cx_excel
                  MESSAGE ID zbc_cx_excel=>mc_message_class
                        TYPE zbc_cx_excel=>mc_message_types-error
                      NUMBER zbc_cx_excel=>mc_messages-generic
                        WITH lref_excel_core->get_text( ).
    ENDTRY.

    IF iv_data_has_header EQ abap_true.
      ASSIGN lr_data->* TO <lt_data>.
      IF <lt_data> IS ASSIGNED.
        IF lines( <lt_data> ) GE 1.
          DELETE <lt_data> INDEX 1 .
        ENDIF.
      ENDIF.
    ENDIF.

    map_fields(
      EXPORTING
        ir_xlsx_data = lr_data
      CHANGING
        cr_itab_data = cr_data ).

  ENDMETHOD.


  METHOD get_data_by_worksheet_name.

    DATA: lr_xl   TYPE REF TO cl_fdt_xl_spreadsheet,
          lr_doc  TYPE REF TO if_fdt_doc_spreadsheet,

          lr_data TYPE REF TO data.

    FIELD-SYMBOLS: <lt_data>  TYPE ANY TABLE,
                   <lv_value> TYPE any.
    TRY.
        CREATE OBJECT lr_xl
          EXPORTING
            document_name = ''
            xdocument     = mv_xml.

        lr_doc ?= lr_xl.

        lr_data = lr_doc->get_itab_from_worksheet( iv_worksheet_name ).

      CATCH cx_fdt_excel_core INTO DATA(lref_excel_core).
    ENDTRY.

  ENDMETHOD.


  METHOD get_xml.
    rv_xml = mv_xml.
  ENDMETHOD.


  METHOD map_fields.

    FIELD-SYMBOLS: <lt_xlsx_data> TYPE table.
    FIELD-SYMBOLS: <lt_itab_data> TYPE table.

    DATA: ls_components      TYPE abap_compdescr,
          lo_xlsx_strucdescr TYPE REF TO cl_abap_structdescr,
          lo_itab_strucdescr TYPE REF TO cl_abap_structdescr,
          lr_xlsx_line       TYPE REF TO data,
          lr_itab_line       TYPE REF TO data.

    ASSIGN ir_xlsx_data->* TO <lt_xlsx_data>.
    ASSIGN cr_itab_data->* TO <lt_itab_data>.

    CREATE DATA lr_xlsx_line LIKE LINE OF <lt_xlsx_data>.
    CREATE DATA lr_itab_line LIKE LINE OF <lt_itab_data>.

    lo_xlsx_strucdescr ?= cl_abap_typedescr=>describe_by_data( lr_xlsx_line->* ).
    lo_itab_strucdescr ?= cl_abap_typedescr=>describe_by_data( lr_itab_line->* ).

    LOOP AT <lt_xlsx_data> ASSIGNING FIELD-SYMBOL(<ls_xlsx_data>).
      APPEND INITIAL LINE TO <lt_itab_data> ASSIGNING FIELD-SYMBOL(<ls_itab_data>).
      LOOP AT lo_xlsx_strucdescr->components ASSIGNING FIELD-SYMBOL(<ls_xlsx_component>).
        READ TABLE lo_itab_strucdescr->components ASSIGNING FIELD-SYMBOL(<ls_itab_component>) INDEX sy-tabix.

        ASSIGN COMPONENT <ls_xlsx_component>-name OF STRUCTURE <ls_xlsx_data> TO FIELD-SYMBOL(<lv_xlsx_value>).
        ASSIGN COMPONENT <ls_itab_component>-name OF STRUCTURE <ls_itab_data> TO FIELD-SYMBOL(<lv_itab_value>).

        IF <lv_xlsx_value> IS INITIAL.
          CONTINUE.
        ENDIF.

        CASE <ls_itab_component>-type_kind.
          WHEN 'D'.
            <lv_itab_value> = |{ <lv_xlsx_value>(4) }{ <lv_xlsx_value>+5(2) }{ <lv_xlsx_value>+8(2) }|.

          WHEN 'C'.
            <lv_itab_value> = <lv_xlsx_value>.

          WHEN OTHERS.

            "Excel tipi Float mı?
            IF <lv_xlsx_value> CP '*E*'.
              DATA(lv_float) = VALUE float( ).
              TRY .
                  lv_float = <lv_xlsx_value>.
                  <lv_itab_value> = CONV #( lv_float ).
                  CONTINUE.
                CATCH cx_sy_conversion_no_number INTO DATA(lo_conv_no_number).
              ENDTRY.
            ENDIF.

            <lv_itab_value> = <lv_xlsx_value>.
        ENDCASE.

      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.


  METHOD upload_xlsx.

    DATA: lt_records  TYPE solix_tab.

    ro_myself = NEW #( ).

    IF iv_filename IS INITIAL.
      ev_filepath = file_open_dialog( ).
      IF ev_filepath IS INITIAL.
        RETURN.
      ENDIF.

    ELSE.
      IF abap_true NE check_filename( iv_filename = iv_filename ).
        "File extension must be XLSX!
        RAISE EXCEPTION TYPE zbc_cx_excel
                  MESSAGE ID zbc_cx_excel=>mc_message_class
                        TYPE zbc_cx_excel=>mc_message_types-error
                      NUMBER zbc_cx_excel=>mc_messages-file_ext_must_be_xlsx.
      ENDIF.

      ev_filepath = iv_filename.
    ENDIF.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename                = ev_filepath
        filetype                = 'BIN'
      IMPORTING
        filelength              = DATA(lv_filelength)
      CHANGING
        data_tab                = lt_records
      EXCEPTIONS
        file_open_error         = 1  " File does not exist and cannot be opened
        file_read_error         = 2  " Error when reading file
        no_batch                = 3  " Cannot execute front-end function in background
        gui_refuse_filetransfer = 4  " Incorrect front end or error on front end
        invalid_type            = 5  " Incorrect parameter FILETYPE
        no_authority            = 6  " No upload authorization
        unknown_error           = 7  " Unknown error
        bad_data_format         = 8  " Cannot Interpret Data in File
        header_not_allowed      = 9  " Invalid header
        separator_not_allowed   = 10 " Invalid separator
        header_too_long         = 11 " Header information currently restricted to 1023 bytes
        unknown_dp_error        = 12 " Error when calling data provider
        access_denied           = 13 " Access to file denied.
        dp_out_of_memory        = 14 " Not enough memory in data provider
        disk_full               = 15 " Storage medium is full.
        dp_timeout              = 16 " Data provider timeout
        not_supported_by_gui    = 17 " GUI does not support this
        error_no_gui            = 18 " GUI not available
        OTHERS                  = 19
    ).

    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 INTO DATA(lv_gui_upload_message).

      "& & & &
      RAISE EXCEPTION TYPE zbc_cx_excel
                MESSAGE ID zbc_cx_excel=>mc_message_class
                      TYPE zbc_cx_excel=>mc_message_types-error
                    NUMBER zbc_cx_excel=>mc_messages-generic
                      WITH lv_gui_upload_message.
    ENDIF.

    ro_myself->mv_xml = cl_bcs_convert=>solix_to_xstring( it_solix = lt_records ).

  ENDMETHOD.


  METHOD check_reference_is_table.

    DATA: lo_type_def    TYPE REF TO cl_abap_typedescr.
    DATA: lo_table_def   TYPE REF TO cl_abap_tabledescr.

    rv_bool = abap_true.

    TRY .
        lo_type_def  = cl_abap_tabledescr=>describe_by_data_ref( ir_data ).
        lo_table_def ?= lo_type_def.
      CATCH cx_sy_move_cast_error INTO DATA(lo_cast_error).
        rv_bool = abap_false.
    ENDTRY.

  ENDMETHOD.


  METHOD check_filename.

    rv_bool = abap_false.

    DATA(lv_splitter) = VALUE char1( ).

    IF iv_filename CA '/'.
      lv_splitter = '/'.
    ELSE.
      lv_splitter = '\'.
    ENDIF.

    SPLIT iv_filename AT lv_splitter INTO TABLE DATA(lt_directories).

    DATA(lv_filename) = VALUE #( lt_directories[ lines( lt_directories ) ] OPTIONAL ).

    SPLIT lv_filename AT '.' INTO DATA(lv_name) DATA(lv_extension).

    TRANSLATE lv_extension TO UPPER CASE.

    IF lv_extension EQ 'XLSX'.
      rv_bool = abap_true.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
