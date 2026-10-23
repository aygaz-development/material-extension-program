*&---------------------------------------------------------------------*
*& Include zaygmmgen_p0039_cls
*&---------------------------------------------------------------------*

CLASS lcl_program DEFINITION CREATE PRIVATE.

  PUBLIC SECTION.

    TYPES: ty_creation_type TYPE text10.

    CONSTANTS: BEGIN OF sc_containers,
                 material         TYPE char30 VALUE 'CONTAINER_MATERIAL',
                 werks_source     TYPE char30 VALUE 'CONTAINER_WERKS_SOURCE',
                 werks_target     TYPE char30 VALUE 'CONTAINER_WERKS_TARGET',
                 vkorg_source     TYPE char30 VALUE 'CONTAINER_VKORG_SOURCE',
                 vkorg_target     TYPE char30 VALUE 'CONTAINER_VKORG_TARGET',
                 valuation_source TYPE char30 VALUE 'CONTAINER_VALUATION_SOURCE',
                 valuation_target TYPE char30 VALUE 'CONTAINER_VALUATION_TARGET',
               END OF sc_containers,

               BEGIN OF sc_creation_types,
                 manuel TYPE ty_creation_type VALUE 'Manuel',
                 excel  TYPE ty_creation_type VALUE 'Excel',
               END OF sc_creation_types,

               BEGIN OF sc_function_codes,
                 back                     TYPE string VALUE 'BACK',
                 cancel                   TYPE string VALUE 'CANCEL',
                 exit                     TYPE string VALUE 'EXIT',

                 manual                   TYPE string VALUE 'MANUAL',
                 download_excel_template  TYPE string VALUE 'DOWN_EXCEL',
                 upload_excel             TYPE string VALUE 'UPLOAD_EXC',
                 create                   TYPE string VALUE 'CREATE',

                 assign_werks             TYPE string VALUE 'ASSIGN_WERKS',
                 assign_vkorg             TYPE string VALUE 'ASSIGN_VKORG',
                 assign_valuation         TYPE string VALUE 'ASSIGN_VALUATION',

                 unassign_werks           TYPE string VALUE 'UNASSIGN_WERKS',
                 unassign_vkorg           TYPE string VALUE 'UNASSIGN_VKORG',
                 unassign_valuation       TYPE string VALUE 'UNASSIGN_VALUATION',

                 okay                     TYPE string VALUE 'OKAY',
                 material_number_genarate TYPE string VALUE 'MATNR_GENARATE',

                 plant_copy               TYPE string VALUE 'PLANT_COPY',
                 extended_search          TYPE string VALUE 'EXTENDED_SEARCH',
               END OF sc_function_codes,

               BEGIN OF sc_structures,
                 material  TYPE tabname VALUE 'ZAYGMM_S001',
                 werks     TYPE tabname VALUE 'ZAYGMM_S002',
                 vkorg     TYPE tabname VALUE 'ZAYGMM_S003',
                 valuation TYPE tabname VALUE 'ZAYGMM_S004',
                 excel     TYPE tabname VALUE 'ZAYGMM_S005',
               END OF sc_structures.

    CLASS-METHODS:
      get_instance
        RETURNING
          VALUE(ro_object) TYPE REF TO lcl_program.

    METHODS:
      constructor,
      start_of_selection,

      status_0100,
      user_command_0100,

      status_0200,
      user_command_0200,
      pbo_field_matnr,

      status_0300,
      user_command_0300,
      user_command_0300_at_exit,

      double_click
        FOR EVENT double_click OF cl_gui_alv_grid
        IMPORTING
          e_row
          e_column
          sender,

      hotspot_click
        FOR EVENT hotspot_click OF cl_gui_alv_grid
        IMPORTING
          e_row_id
          e_column_id
          es_row_no,

      plant_extension
        IMPORTING
          VALUE(iv_matnr)        TYPE mara-matnr OPTIONAL
          VALUE(iv_werks_source) TYPE marc-werks OPTIONAL
          VALUE(iv_werks_target) TYPE marc-werks OPTIONAL
        EXPORTING
          VALUE(et_messages)     TYPE bapirettab.

  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES: ty_material TYPE zaygmm_s001,
           tt_material TYPE TABLE OF ty_material WITH EMPTY KEY,

           ty_excel    TYPE zaygmm_s005,
           tt_excel    TYPE TABLE OF ty_excel WITH EMPTY KEY,

           BEGIN OF ty_lgnum,
             werks TYPE werks_d,
             lgnum TYPE lgnum,
           END OF ty_lgnum,
           tt_lgnum TYPE TABLE OF ty_lgnum WITH EMPTY KEY,

           BEGIN OF ty_lgort,
             matnr TYPE matnr,
             werks TYPE werks_d,
             lgort TYPE lgort_d,
           END OF ty_lgort,
           tt_lgort        TYPE TABLE OF ty_lgort WITH EMPTY KEY,

           ty_werks        TYPE zaygmm_s002,
           tt_werks        TYPE SORTED TABLE OF ty_werks WITH UNIQUE KEY matnr werks,
           tt_werks_ui     TYPE TABLE OF ty_werks WITH EMPTY KEY,

           ty_vkorg        TYPE zaygmm_s003,
           tt_vkorg        TYPE SORTED TABLE OF ty_vkorg WITH UNIQUE KEY matnr vkorg vtweg,
           tt_vkorg_ui     TYPE TABLE OF ty_vkorg WITH EMPTY KEY,

           ty_valuation    TYPE zaygmm_s004,
           tt_valuation    TYPE SORTED TABLE OF ty_valuation WITH UNIQUE KEY matnr bwkey bwtar,
           tt_valuation_ui TYPE TABLE OF ty_valuation WITH EMPTY KEY.

    CONSTANTS: BEGIN OF sc_screens,
                 main          TYPE char4 VALUE '0100',
                 manuel_create TYPE char4 VALUE '0200',
                 plant_copy    TYPE char4 VALUE '0300',
               END OF sc_screens.

    CLASS-DATA: so_instance    TYPE REF TO lcl_program.

    DATA: mo_container_material         TYPE REF TO cl_gui_custom_container,

          mo_container_werks_source     TYPE REF TO cl_gui_custom_container,
          mo_container_werks_target     TYPE REF TO cl_gui_custom_container,

          mo_container_vkorg_source     TYPE REF TO cl_gui_custom_container,
          mo_container_vkorg_target     TYPE REF TO cl_gui_custom_container,

          mo_container_valuation_source TYPE REF TO cl_gui_custom_container,
          mo_container_valuation_target TYPE REF TO cl_gui_custom_container,

          mo_alv_material               TYPE REF TO cl_gui_alv_grid,

          mo_alv_werks_source           TYPE REF TO cl_gui_alv_grid,
          mo_alv_werks_target           TYPE REF TO cl_gui_alv_grid,

          mo_alv_vkorg_source           TYPE REF TO cl_gui_alv_grid,
          mo_alv_vkorg_target           TYPE REF TO cl_gui_alv_grid,

          mo_alv_valuation_source       TYPE REF TO cl_gui_alv_grid,
          mo_alv_valuation_target       TYPE REF TO cl_gui_alv_grid,

          mt_material                   TYPE tt_material,
          mt_werks                      TYPE tt_werks,
          mt_lgort                      TYPE tt_lgort,
          mt_vkorg                      TYPE tt_vkorg,
          mt_valuation                  TYPE tt_valuation,

          mr_material_selected          TYPE REF TO ty_material,

          mt_werks_source               TYPE tt_werks_ui,
          mt_werks_target               TYPE tt_werks_ui,

          mt_vkorg_source               TYPE tt_vkorg_ui,
          mt_vkorg_target               TYPE tt_vkorg_ui,

          mt_valuation_source           TYPE tt_valuation_ui,
          mt_valuation_target           TYPE tt_valuation_ui,

          mv_draft_count                TYPE i,

          mt_lgnum                      TYPE tt_lgnum.

    METHODS: build_alv,
      get_field_catalog
        IMPORTING
          iv_structure_name       TYPE tabname
        RETURNING
          VALUE(rt_field_catalog) TYPE lvc_t_fcat,
      excel_upload,
      download_excel_template,
      get_layout
        RETURNING
          VALUE(rs_layout) TYPE lvc_s_layo,

      determine_material_extension,

      assign_selected_werks,
      assign_selected_vkorg,
      assign_selected_valuation,

      unassign_selected_werks,
      unassign_selected_vkorg,
      unassign_selected_valuation,

      create_and_extend_material
        IMPORTING
          iv_matnr_source TYPE matnr
          iv_matnr        TYPE matnr
        EXPORTING
          et_return       TYPE bapirettab,

      material_get_number_next
        IMPORTING
          iv_mtart        TYPE mtart
        RETURNING
          VALUE(rv_matnr) TYPE matnr,

      create,

      fill_bapix
        IMPORTING
          ir_data TYPE REF TO data,

      refresh_display,

      popup_get_values
        IMPORTING
          iv_no_value_check    TYPE flag DEFAULT abap_false
          iv_popup_title       TYPE text50
          iv_start_column      TYPE char3 DEFAULT '5'
          iv_start_row         TYPE char3 DEFAULT '5'
        EXPORTING
          VALUE(ev_returncode) TYPE char1
        CHANGING
          ct_fields            TYPE ty_sval,

      ask_werks_fields
        IMPORTING
          iv_werks   TYPE marc-werks
          iv_minbe   TYPE marc-minbe OPTIONAL
          iv_bstrf   TYPE marc-bstrf OPTIONAL
          iv_dismm   TYPE marc-dismm OPTIONAL
          iv_lgnum   TYPE mlgn-lgnum OPTIONAL
        EXPORTING
          ev_aborted TYPE abap_bool
          ev_minbe   TYPE marc-minbe
          ev_bstrf   TYPE marc-bstrf
          ev_dismm   TYPE marc-dismm,

      ask_vkorg_fields
        IMPORTING
          iv_vkorg     TYPE mvke-vkorg
          iv_vtweg     TYPE mvke-vtweg OPTIONAL
          iv_kondm     TYPE mvke-kondm OPTIONAL
          iv_ktgrm     TYPE mvke-ktgrm OPTIONAL
          iv_mtpos     TYPE mvke-mtpos OPTIONAL
          iv_tatyp     TYPE tskm-tatyp OPTIONAL
          iv_taxkm     TYPE tskm-taxkm OPTIONAL
          iv_taxim     TYPE mlan-taxim OPTIONAL
          iv_ask_taxim TYPE abap_bool
        EXPORTING
          ev_aborted   TYPE abap_bool
          ev_kondm     TYPE mvke-kondm
          ev_ktgrm     TYPE mvke-ktgrm
          ev_mtpos     TYPE mvke-mtpos
          ev_tatyp     TYPE tskm-tatyp
          ev_taxkm     TYPE tskm-taxkm
          ev_taxim     TYPE mlan-taxim,

      ask_valuation_fields
        IMPORTING
          iv_bwkey   TYPE mbew-bwkey
          iv_bwtar   TYPE mbew-bwtar
          iv_stprs   TYPE mbew-stprs OPTIONAL
        EXPORTING
          ev_aborted TYPE abap_bool
          ev_stprs   TYPE mbew-stprs,

      get_reference_material_data
        IMPORTING
          iv_matnr               TYPE matnr
          iv_plant               TYPE werks_d OPTIONAL
          iv_salesorg            TYPE vkorg OPTIONAL
          iv_distr_chan          TYPE vtweg OPTIONAL
          iv_val_area            TYPE bwkey OPTIONAL
          iv_val_type            TYPE bwtar_d OPTIONAL
          iv_whse_no             TYPE lgnum OPTIONAL
        EXPORTING
          es_plantdata           TYPE bapi_marc_ga
          es_warehousenumberdata TYPE bapi_mlgn_ga
          es_salesdata           TYPE bapi_mvke_ga
          es_valuationdata       TYPE bapi_mbew_ga
          es_taxclassification   TYPE bapi_mlan_ga
          ev_taxim               TYPE mlan-taxim.

ENDCLASS.

CLASS lcl_program IMPLEMENTATION.

  METHOD get_instance.
    IF so_instance IS BOUND.
      ro_object = so_instance.
      RETURN.
    ENDIF.
    ro_object = so_instance = NEW #( ).
  ENDMETHOD.

  METHOD constructor.
    mv_draft_count = 1.

    SELECT werks,
           lgnum
      FROM t320
     WHERE werks IN ('1111', '1112', '1114', '1911' )
       AND lgort IN ( SELECT MAX( lgort ) FROM t320 AS sub WHERE werks EQ t320~werks )
      INTO TABLE @mt_lgnum.
  ENDMETHOD.

  METHOD start_of_selection.
    CALL SCREEN sc_screens-main.
  ENDMETHOD.

  METHOD status_0100.

    DATA lt_fcode TYPE TABLE OF sy-ucomm.

    CLEAR lt_fcode.
    IF mt_material IS NOT INITIAL.
      lt_fcode = VALUE #( ( CONV #( sc_function_codes-upload_excel ) )
                          ( CONV #( sc_function_codes-manual ) ) ).

    ENDIF.

    SET PF-STATUS sc_screens-main EXCLUDING lt_fcode.
    build_alv( ).
  ENDMETHOD.

  METHOD user_command_0100.

    CASE sy-ucomm.
      WHEN sc_function_codes-back   OR
           sc_function_codes-cancel OR
           sc_function_codes-exit.

        LEAVE PROGRAM.

      WHEN sc_function_codes-manual.
        CALL SCREEN sc_screens-manuel_create STARTING AT 30 5.
        determine_material_extension( ).
        refresh_display( ).

      WHEN sc_function_codes-download_excel_template.
        download_excel_template( ).

      WHEN sc_function_codes-upload_excel.
        excel_upload( ).
        determine_material_extension( ).
        refresh_display( ).

      WHEN sc_function_codes-assign_werks.
        assign_selected_werks( ).
        refresh_display( ).

      WHEN sc_function_codes-unassign_werks.
        unassign_selected_werks( ).
        refresh_display( ).

      WHEN sc_function_codes-assign_vkorg.
        assign_selected_vkorg( ).
        refresh_display( ).

      WHEN sc_function_codes-unassign_vkorg.
        unassign_selected_vkorg( ).
        refresh_display( ).

      WHEN sc_function_codes-assign_valuation.
        assign_selected_valuation( ).
        refresh_display( ).

      WHEN sc_function_codes-unassign_valuation.
        unassign_selected_valuation( ).
        refresh_display( ).

      WHEN sc_function_codes-create.
        create( ).
        refresh_display( ).

      WHEN sc_function_codes-plant_copy.
        CALL SCREEN sc_screens-plant_copy STARTING AT 30 5.

    ENDCASE.

  ENDMETHOD.

  METHOD status_0200.
    SET PF-STATUS sc_screens-manuel_create.
  ENDMETHOD.

  METHOD user_command_0200.

    CASE sy-ucomm.
      WHEN sc_function_codes-back   OR
           sc_function_codes-cancel OR
           sc_function_codes-exit.

        CLEAR gs_manuel_create_scr.
        LEAVE TO SCREEN 0.
    ENDCASE.

    CASE sy-ucomm.
      WHEN sc_function_codes-okay.

        pbo_field_matnr( ).

        APPEND gs_manuel_create_scr TO mt_material ASSIGNING FIELD-SYMBOL(<ls_material>).
        <ls_material>-icon          = icon_led_inactive.
        <ls_material>-creation_type = sc_creation_types-manuel.
        CLEAR gs_manuel_create_scr.
        LEAVE TO SCREEN 0.

      WHEN sc_function_codes-material_number_genarate.

        gs_manuel_create_scr-matnr_target = material_get_number_next( gs_manuel_create_scr-mtart_source ).

    ENDCASE.

  ENDMETHOD.

  METHOD pbo_field_matnr.

    IF gs_manuel_create_scr-matnr_source IS NOT INITIAL.

      IF gs_manuel_create_scr-maktx_source IS INITIAL.

        SELECT SINGLE makt~maktx, mara~mtart, t134t~mtbez, mara~zziy

          FROM mara

          LEFT JOIN makt ON makt~matnr EQ mara~matnr
                        AND makt~spras EQ @sy-langu

          JOIN t134t ON t134t~mtart EQ mara~mtart
                    AND t134t~spras EQ @sy-langu

         WHERE makt~matnr EQ @gs_manuel_create_scr-matnr_source

          INTO ( @gs_manuel_create_scr-maktx_source, @gs_manuel_create_scr-mtart_source, @gs_manuel_create_scr-mtbez_source , @gs_manuel_create_scr-zziy ).

        SELECT SINGLE makt~maktx

          FROM mara

          LEFT JOIN makt ON makt~matnr EQ mara~matnr
                        AND makt~spras EQ @sy-langu

         WHERE makt~matnr EQ @gs_manuel_create_scr-matnr_target

          INTO @gs_manuel_create_scr-maktx_target.

        get_reference_material_data(
          EXPORTING
            iv_matnr             = gs_manuel_create_scr-matnr_source
          IMPORTING
            es_plantdata         = DATA(ls_plantdata)
            es_salesdata         = DATA(ls_salesdata)
            es_valuationdata     = DATA(ls_valuationdata)
            es_taxclassification = DATA(ls_taxclassification)
            ev_taxim             = DATA(lv_taxim) ).

        gs_manuel_create_scr-minbe = ls_plantdata-reorder_pt.
        gs_manuel_create_scr-bstrf = ls_plantdata-round_val.
        gs_manuel_create_scr-dismm = ls_plantdata-mrp_type.

        gs_manuel_create_scr-tatyp = ls_taxclassification-tax_type_1.
        gs_manuel_create_scr-taxkm = ls_taxclassification-taxclass_1.

        gs_manuel_create_scr-kondm = ls_salesdata-mat_pr_grp.
        gs_manuel_create_scr-ktgrm = ls_salesdata-acct_assgt.
        gs_manuel_create_scr-mtpos = ls_salesdata-item_cat.

        gs_manuel_create_scr-taxim = lv_taxim.
        gs_manuel_create_scr-stprs = ls_valuationdata-std_price.

      ENDIF.

      LOOP AT SCREEN.
        IF screen-name EQ 'GS_MANUEL_CREATE_SCR-MATNR_SOURCE'.
          screen-input = 0.
          MODIFY SCREEN.
        ENDIF.
      ENDLOOP.

    ENDIF.

  ENDMETHOD.

  METHOD status_0300.
    SET PF-STATUS sc_screens-manuel_create.
  ENDMETHOD.

  METHOD user_command_0300.
    CASE sy-ucomm.
      WHEN sc_function_codes-back   OR
           sc_function_codes-cancel OR
           sc_function_codes-exit.

        LEAVE TO SCREEN 0.
    ENDCASE.

    CASE sy-ucomm.
      WHEN sc_function_codes-okay.

        CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
          EXPORTING
            input        = gs_plant_copy_scr-matnr
          IMPORTING
            output       = gs_plant_copy_scr-matnr
          EXCEPTIONS
            length_error = 1
            OTHERS       = 2.

        APPEND VALUE #( sign   = 'I'
                        option = 'EQ'
                        low    = gs_plant_copy_scr-matnr ) TO gs_plant_copy_scr-matnr_r.

        SELECT matnr
          FROM mara
         WHERE matnr IN @gs_plant_copy_scr-matnr_r
          INTO TABLE @DATA(lt_mara).

        DATA(lt_messages_all) = VALUE bapirettab( ).
        LOOP AT lt_mara ASSIGNING FIELD-SYMBOL(<ls_mara>).

          plant_extension(
            EXPORTING
              iv_matnr        = <ls_mara>-matnr
              iv_werks_source = gs_plant_copy_scr-werks_source
              iv_werks_target = gs_plant_copy_scr-werks_target
            IMPORTING
              et_messages     = DATA(lt_messages) ).

          APPEND LINES OF lt_messages TO lt_messages_all.

        ENDLOOP.

        IF lt_messages_all IS INITIAL.
          APPEND INITIAL LINE TO lt_messages_all ASSIGNING FIELD-SYMBOL(<ls_message>).
          <ls_message>-id         = '00'.
          <ls_message>-type       = 'W'.
          <ls_message>-number     = '001'.
          <ls_message>-message_v1 = 'İşlem Yapılamadı'.
        ENDIF.

        cl_rmsl_message=>display( lt_messages_all ).

        CLEAR gs_plant_copy_scr.

      WHEN sc_function_codes-extended_search.

        IF gs_plant_copy_scr-matnr   IS NOT INITIAL AND
           gs_plant_copy_scr-matnr_r IS INITIAL.

          gs_plant_copy_scr-matnr_r = VALUE #( ( sign = 'I' option = 'EQ' low = gs_plant_copy_scr-matnr ) ).
        ENDIF.

        CALL FUNCTION 'COMPLEX_SELECTIONS_DIALOG'
          EXPORTING
            title             = 'Select Multiple Value'
            text              = 'Finish Group'
            signed            = 'X'
            lower_case        = ' '
            no_interval_check = 'X'
            just_display      = ' '
            just_incl         = 'X'
            search_help       = 'MAT1'
          TABLES
            range             = gs_plant_copy_scr-matnr_r
          EXCEPTIONS
            no_range_tab      = 1
            cancelled         = 2
            internal_error    = 3
            OTHERS            = 4.

        IF sy-subrc = 0.
          gs_plant_copy_scr-matnr = VALUE #( gs_plant_copy_scr-matnr_r[ 1 ]-low ).
        ENDIF.

    ENDCASE.
  ENDMETHOD.

  METHOD user_command_0300_at_exit.
    CASE sy-ucomm.
      WHEN sc_function_codes-back   OR
           sc_function_codes-cancel OR
           sc_function_codes-exit.

        LEAVE TO SCREEN 0.
    ENDCASE.
  ENDMETHOD.

  METHOD build_alv.

    IF mo_container_material IS BOUND.
      RETURN.
    ENDIF.

    DATA(lt_exclude) = VALUE ui_functions(
*    ( cl_gui_alv_grid=>mc_fc_detail )
*    ( cl_gui_alv_grid=>mc_fc_sort )
*    ( cl_gui_alv_grid=>mc_fc_sort_asc )
*    ( cl_gui_alv_grid=>mc_fc_sort_dsc )
*    ( cl_gui_alv_grid=>mc_fc_find )
*    ( cl_gui_alv_grid=>mc_fc_find_more )
*    ( cl_gui_alv_grid=>mc_fc_filter )
*    ( cl_gui_alv_grid=>mc_fc_print )
    ( cl_gui_alv_grid=>mc_fc_excl_all ) ).

    mo_container_material = NEW cl_gui_custom_container( sc_containers-material ).

    mo_container_werks_source = NEW cl_gui_custom_container( sc_containers-werks_source ).
    mo_container_werks_target = NEW cl_gui_custom_container( sc_containers-werks_target ).

    mo_container_vkorg_source = NEW cl_gui_custom_container( sc_containers-vkorg_source ).
    mo_container_vkorg_target = NEW cl_gui_custom_container( sc_containers-vkorg_target ).

    mo_container_valuation_source = NEW cl_gui_custom_container( sc_containers-valuation_source ).
    mo_container_valuation_target = NEW cl_gui_custom_container( sc_containers-valuation_target ).

    mo_alv_material = NEW cl_gui_alv_grid( mo_container_material ).

    mo_alv_werks_source = NEW cl_gui_alv_grid( mo_container_werks_source ).
    mo_alv_werks_target = NEW cl_gui_alv_grid( mo_container_werks_target ).

    mo_alv_vkorg_source = NEW cl_gui_alv_grid( mo_container_vkorg_source ).
    mo_alv_vkorg_target = NEW cl_gui_alv_grid( mo_container_vkorg_target ).

    mo_alv_valuation_source = NEW cl_gui_alv_grid( mo_container_valuation_source ).
    mo_alv_valuation_target = NEW cl_gui_alv_grid( mo_container_valuation_target ).

    DATA(lt_field_catalog_material)  = get_field_catalog( sc_structures-material ).
    LOOP AT lt_field_catalog_material ASSIGNING FIELD-SYMBOL(<ls_field_catalog>).
      CASE <ls_field_catalog>-fieldname.
        WHEN 'ICON'.
          <ls_field_catalog>-hotspot = abap_true.
      ENDCASE.
    ENDLOOP.

    DATA(lt_field_catalog_werks_t) = get_field_catalog( sc_structures-werks ).
    DATA(lt_field_catalog_werks_s) = lt_field_catalog_werks_t.

    LOOP AT lt_field_catalog_werks_s ASSIGNING <ls_field_catalog>.
      CASE <ls_field_catalog>-fieldname.
        WHEN 'MATNR'.
        WHEN 'WERKS'.
        WHEN OTHERS.
          <ls_field_catalog>-no_out = abap_true.
      ENDCASE.
    ENDLOOP.

    DATA(lt_field_catalog_vkorg_t) = get_field_catalog( sc_structures-vkorg ).
    DATA(lt_field_catalog_vkorg_s) = lt_field_catalog_vkorg_t.

    LOOP AT lt_field_catalog_vkorg_s ASSIGNING <ls_field_catalog>.
      CASE <ls_field_catalog>-fieldname.
        WHEN 'MATNR'.
        WHEN 'VKORG'.
        WHEN 'VTWEG'.
        WHEN OTHERS.
          <ls_field_catalog>-no_out = abap_true.
      ENDCASE.
    ENDLOOP.

    DATA(lt_field_catalog_valuation_t) = get_field_catalog( sc_structures-valuation ).
    DATA(lt_field_catalog_valuation_s) = lt_field_catalog_valuation_t.

    LOOP AT lt_field_catalog_valuation_s ASSIGNING <ls_field_catalog>.
      CASE <ls_field_catalog>-fieldname.
        WHEN 'MATNR'.
        WHEN 'BWKEY'.
        WHEN 'BWTAR'.
        WHEN OTHERS.
          <ls_field_catalog>-no_out = abap_true.
      ENDCASE.
    ENDLOOP.

    "MATERIAL

    SET HANDLER double_click FOR mo_alv_material.
    SET HANDLER hotspot_click FOR mo_alv_material.

    mo_alv_material->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
      CHANGING
        it_outtab                     = mt_material
        it_fieldcatalog               = lt_field_catalog_material
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "WERKS SOURCE
    mo_alv_werks_source->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_werks_source
        it_fieldcatalog               = lt_field_catalog_werks_s
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "WERKS TARGET
    mo_alv_werks_target->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_werks_target
        it_fieldcatalog               = lt_field_catalog_werks_t
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "VKORG SOURCE
    mo_alv_vkorg_source->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_vkorg_source
        it_fieldcatalog               = lt_field_catalog_vkorg_s
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "VKORG TARGET
    mo_alv_vkorg_target->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_vkorg_target
        it_fieldcatalog               = lt_field_catalog_vkorg_t
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "VALUATION SOURCE
    mo_alv_valuation_source->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_valuation_source
        it_fieldcatalog               = lt_field_catalog_valuation_s
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "VALUATION TARGET
    mo_alv_valuation_target->set_table_for_first_display(
      EXPORTING
        is_layout                     = get_layout( )
        it_toolbar_excluding          = lt_exclude
      CHANGING
        it_outtab                     = mt_valuation_target
        it_fieldcatalog               = lt_field_catalog_valuation_t
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

  ENDMETHOD.

  METHOD double_click.

    READ TABLE mt_material ASSIGNING FIELD-SYMBOL(<ls_material>) INDEX e_row.
    IF sy-subrc IS INITIAL.
      mr_material_selected = REF #( <ls_material> ).
      refresh_display( ).
    ENDIF.

  ENDMETHOD.


  METHOD hotspot_click.
    READ TABLE mt_material ASSIGNING FIELD-SYMBOL(<ls_material>) INDEX e_row_id-index.
    IF sy-subrc IS INITIAL.
      cl_rmsl_message=>display( <ls_material>-messages ).
    ENDIF.
  ENDMETHOD.

  METHOD refresh_display.

    IF mr_material_selected IS BOUND AND mr_material_selected->* IS NOT INITIAL.

      mt_werks_target     = VALUE #( FOR wa_werks IN mt_werks WHERE ( matnr = mr_material_selected->matnr_target ) ( wa_werks ) ).
      mt_vkorg_target     = VALUE #( FOR wa_vkorg IN mt_vkorg WHERE ( matnr = mr_material_selected->matnr_target ) ( wa_vkorg ) ).
      mt_valuation_target = VALUE #( FOR wa_valuation IN mt_valuation WHERE ( matnr = mr_material_selected->matnr_target ) ( wa_valuation ) ).

      CLEAR mt_werks_source.
      LOOP AT mt_werks ASSIGNING FIELD-SYMBOL(<ls_werks>) WHERE matnr EQ mr_material_selected->matnr_source.
        READ TABLE mt_werks_target TRANSPORTING NO FIELDS WITH KEY matnr = mr_material_selected->matnr_target
                                                                   werks = <ls_werks>-werks.
        IF sy-subrc IS NOT INITIAL.
          APPEND <ls_werks> TO mt_werks_source.
        ENDIF.
      ENDLOOP.

      CLEAR mt_vkorg_source.
      LOOP AT mt_vkorg ASSIGNING FIELD-SYMBOL(<ls_vkorg>) WHERE matnr EQ mr_material_selected->matnr_source.
        READ TABLE mt_vkorg_target TRANSPORTING NO FIELDS WITH KEY matnr = mr_material_selected->matnr_target
                                                                   vkorg = <ls_vkorg>-vkorg
                                                                   vtweg = <ls_vkorg>-vtweg.
        IF sy-subrc IS NOT INITIAL.
          APPEND <ls_vkorg> TO mt_vkorg_source.
        ENDIF.
      ENDLOOP.

      CLEAR mt_valuation_source.
      LOOP AT mt_valuation ASSIGNING FIELD-SYMBOL(<ls_valuation>) WHERE matnr EQ mr_material_selected->matnr_source.
        READ TABLE mt_valuation_target TRANSPORTING NO FIELDS WITH KEY matnr = mr_material_selected->matnr_target
                                                                       bwkey = <ls_valuation>-bwkey
                                                                       bwtar = <ls_valuation>-bwtar.
        IF sy-subrc IS NOT INITIAL.
          APPEND <ls_valuation> TO mt_valuation_source.
        ENDIF.
      ENDLOOP.

      mo_alv_werks_source->refresh_table_display( ).
      mo_alv_vkorg_source->refresh_table_display( ).
      mo_alv_valuation_source->refresh_table_display( ).

      mo_alv_werks_target->refresh_table_display( ).
      mo_alv_vkorg_target->refresh_table_display( ).
      mo_alv_valuation_target->refresh_table_display( ).

    ENDIF.

    SELECT material~matnr_target
      FROM @mt_material AS material
      JOIN mara ON mara~matnr EQ material~matnr_target
     ORDER BY material~matnr_target
      INTO TABLE @DATA(lt_created_materials).

    LOOP AT mt_material ASSIGNING FIELD-SYMBOL(<ls_material>).

      IF line_exists( mt_werks[ matnr = <ls_material>-matnr_target ] ) OR
         line_exists( mt_vkorg[ matnr = <ls_material>-matnr_target ] ) OR
         line_exists( mt_valuation[ matnr = <ls_material>-matnr_target ] ).

        <ls_material>-icon = icon_led_yellow.
      ELSE.
        <ls_material>-icon = icon_led_inactive.
      ENDIF.

      READ TABLE lt_created_materials TRANSPORTING NO FIELDS WITH KEY matnr_target = <ls_material>-matnr_target BINARY SEARCH.
      IF sy-subrc IS INITIAL.
        <ls_material>-icon = icon_led_green.
      ENDIF.

      LOOP AT <ls_material>-messages TRANSPORTING NO FIELDS WHERE type CA 'EAX'.
        EXIT.
      ENDLOOP.
      IF sy-subrc IS INITIAL.
        <ls_material>-icon = icon_led_red.
      ENDIF.

    ENDLOOP.

    mo_alv_material->refresh_table_display( i_soft_refresh = abap_true ).

    IF mr_material_selected IS BOUND.
      READ TABLE mt_material TRANSPORTING NO FIELDS WITH KEY matnr_source = mr_material_selected->matnr_source
                                                             matnr_target = mr_material_selected->matnr_target.

      mo_alv_material->set_selected_rows( it_row_no = VALUE #( ( row_id = sy-tabix ) ) ).
    ENDIF.

  ENDMETHOD.

  METHOD get_field_catalog.
    CLEAR rt_field_catalog.
    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
      EXPORTING
        i_structure_name       = iv_structure_name
        i_bypassing_buffer     = abap_true
      CHANGING
        ct_fieldcat            = rt_field_catalog
      EXCEPTIONS
        inconsistent_interface = 1
        program_error          = 2
        OTHERS                 = 3.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDMETHOD.

  METHOD get_layout.
    rs_layout = VALUE lvc_s_layo( sel_mode = 'A' zebra = abap_true col_opt = abap_true cwidth_opt = abap_true ).
  ENDMETHOD.

  METHOD download_excel_template.

    DATA(lr_excel) = NEW tt_excel( ).

    TRY.
        DATA(lo_excel) = zbc_cl_excel=>create_with_internal_table( ir_data          = lr_excel
                                                                   it_field_catalog = get_field_catalog( sc_structures-excel ) ).

        lo_excel->download_xlsx( ).
      CATCH zbc_cx_excel.
    ENDTRY.

  ENDMETHOD.

  METHOD excel_upload.

    DATA lr_excel TYPE REF TO data.

    CREATE DATA lr_excel TYPE tt_excel.

    FIELD-SYMBOLS <fs_excel> TYPE tt_excel.

    ASSIGN lr_excel->* TO <fs_excel>.

    TRY.
        DATA(lo_excel) = zbc_cl_excel=>upload_xlsx( ). "'Z:\Export.xlsx'
        lo_excel->get_data_by_worksheet_index(
          EXPORTING
            iv_index           = 1
            iv_data_has_header = abap_true
          CHANGING
            cr_data            = lr_excel ).
      CATCH zbc_cx_excel.
    ENDTRY.

    mt_material = CORRESPONDING #( BASE ( mt_material ) <fs_excel> ).

    LOOP AT mt_material ASSIGNING FIELD-SYMBOL(<ls_material>).

      CHECK <ls_material>-creation_type IS INITIAL.

      CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
        EXPORTING
          input        = <ls_material>-matnr_source
        IMPORTING
          output       = <ls_material>-matnr_source
        EXCEPTIONS
          length_error = 1
          OTHERS       = 2.

      CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
        EXPORTING
          input        = <ls_material>-matnr_target
        IMPORTING
          output       = <ls_material>-matnr_target
        EXCEPTIONS
          length_error = 1
          OTHERS       = 2.

      <ls_material>-icon          = icon_led_inactive.
      <ls_material>-creation_type = sc_creation_types-excel.
    ENDLOOP.

  ENDMETHOD.

  METHOD determine_material_extension.

    SELECT mara~matnr,
           makt~maktx,
           mara~mtart,
           t134t~mtbez

      FROM @mt_material AS material

      JOIN mara ON mara~matnr EQ material~matnr_source

      LEFT JOIN makt ON makt~matnr EQ mara~matnr
                    AND makt~spras EQ @sy-langu

      LEFT JOIN t134t ON t134t~mtart EQ mara~mtart
                     AND t134t~spras EQ @sy-langu

      INTO TABLE @DATA(lt_mara).

    SELECT mara~matnr,
           makt~maktx

      FROM @mt_material AS material

      JOIN mara ON mara~matnr EQ material~matnr_target

      LEFT JOIN makt ON makt~matnr EQ mara~matnr
                    AND makt~spras EQ @sy-langu

      APPENDING TABLE @lt_mara.

    SELECT DISTINCT material~matnr_source,
           marc~werks
      FROM @mt_material AS material
      JOIN marc ON marc~matnr EQ material~matnr_source
      INTO TABLE @mt_werks.

    SELECT DISTINCT material~matnr_target,
           marc~werks
      FROM @mt_material AS material
      JOIN marc ON marc~matnr EQ material~matnr_target
     APPENDING TABLE @mt_werks.

    LOOP AT mt_werks ASSIGNING FIELD-SYMBOL(<ls_werks>).
      READ TABLE mt_lgnum ASSIGNING FIELD-SYMBOL(<ls_lgnum>) WITH KEY werks = <ls_werks>-werks.
      IF sy-subrc IS INITIAL.
        <ls_werks>-lgnum = <ls_lgnum>-lgnum.
      ENDIF.
    ENDLOOP.

    SELECT DISTINCT material~matnr_source,
           mard~werks,
           mard~lgort
      FROM @mt_material AS material
      JOIN mard ON mard~matnr EQ material~matnr_source
      INTO TABLE @mt_lgort.

    SELECT DISTINCT material~matnr_target,
           mard~werks,
           mard~lgort
      FROM @mt_material AS material
      JOIN mard ON mard~matnr EQ material~matnr_target
      APPENDING TABLE @mt_lgort.

    SELECT DISTINCT material~matnr_source,
           mvke~vkorg,
           mvke~vtweg
      FROM @mt_material AS material
      JOIN mvke ON mvke~matnr EQ material~matnr_source
      INTO TABLE @mt_vkorg.

    SELECT DISTINCT material~matnr_target,
           mvke~vkorg,
           mvke~vtweg
      FROM @mt_material AS material
      JOIN mvke ON mvke~matnr EQ material~matnr_target
      APPENDING TABLE @mt_vkorg.

    SELECT DISTINCT material~matnr_source,
           mbew~bwkey,
           mbew~bwtar
      FROM @mt_material AS material
      JOIN mbew ON mbew~matnr EQ material~matnr_source
      INTO TABLE @mt_valuation.

    SELECT DISTINCT material~matnr_target,
           mbew~bwkey,
           mbew~bwtar
      FROM @mt_material AS material
      JOIN mbew ON mbew~matnr EQ material~matnr_target
     APPENDING TABLE @mt_valuation.

    LOOP AT mt_material ASSIGNING FIELD-SYMBOL(<ls_material>).

*      IF line_exists( mt_werks[ matnr = <ls_material>-matnr_target ] ) OR
*         line_exists( mt_vkorg[ matnr = <ls_material>-matnr_target ] ) OR
*         line_exists( mt_valuation[ matnr = <ls_material>-matnr_target ] ).
*        CONTINUE.
*      ENDIF.

      READ TABLE lt_mara ASSIGNING FIELD-SYMBOL(<ls_mara>) WITH KEY matnr = <ls_material>-matnr_source.
      IF sy-subrc IS INITIAL.
        <ls_material>-maktx_source = <ls_mara>-maktx.
        <ls_material>-mtart_source = <ls_mara>-mtart.
        <ls_material>-mtbez_source = <ls_mara>-mtbez.
      ENDIF.

      IF <ls_material>-matnr_target IS INITIAL.
        <ls_material>-matnr_target = |${ mv_draft_count }|.
        ADD 1 TO mv_draft_count.
      ELSE.
        IF <ls_material>-maktx_target IS INITIAL.
          READ TABLE lt_mara ASSIGNING <ls_mara> WITH KEY matnr = <ls_material>-matnr_target.
          IF sy-subrc IS INITIAL.
            <ls_material>-maktx_target = <ls_mara>-maktx.
          ENDIF.
        ENDIF.
      ENDIF.

      DATA(ls_werks)     = VALUE ty_werks( ).
      DATA(ls_vkorg)     = VALUE ty_vkorg( ).
      DATA(ls_valuation) = VALUE ty_valuation( ).

      CASE <ls_material>-creation_type.
        WHEN sc_creation_types-manuel.

          ls_werks-minbe = <ls_material>-minbe.
          ls_werks-bstrf = <ls_material>-bstrf.
          ls_werks-dismm = <ls_material>-dismm.

          ls_vkorg-kondm = <ls_material>-kondm.
          ls_vkorg-ktgrm = <ls_material>-ktgrm.
          ls_vkorg-mtpos = <ls_material>-mtpos.
          ls_vkorg-tatyp = <ls_material>-tatyp.
          ls_vkorg-taxkm = <ls_material>-taxkm.
          ls_vkorg-taxim = <ls_material>-taxim.

          ls_valuation-stprs = <ls_material>-stprs.

          IF ls_werks IS NOT INITIAL.
            LOOP AT mt_werks ASSIGNING <ls_werks> USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
              IF line_exists( mt_werks[ matnr = <ls_material>-matnr_target
                                        werks = <ls_werks>-werks ] ).
                CONTINUE.
              ENDIF.
              ls_werks-matnr = <ls_material>-matnr_target.
              ls_werks-werks = <ls_werks>-werks.
              ls_werks-lgnum = <ls_werks>-lgnum.
              INSERT ls_werks INTO TABLE mt_werks.
            ENDLOOP.
          ENDIF.

          IF ls_vkorg IS NOT INITIAL.
            LOOP AT mt_vkorg ASSIGNING FIELD-SYMBOL(<ls_vkorg>) USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
              IF line_exists( mt_vkorg[ matnr = <ls_material>-matnr_target
                                        vkorg = <ls_vkorg>-vkorg
                                        vtweg = <ls_vkorg>-vtweg ] ).
                CONTINUE.
              ENDIF.
              ls_vkorg-matnr = <ls_material>-matnr_target.
              ls_vkorg-vkorg = <ls_vkorg>-vkorg.
              ls_vkorg-vtweg = <ls_vkorg>-vtweg.
              INSERT ls_vkorg INTO TABLE mt_vkorg.
            ENDLOOP.
          ENDIF.

          IF ls_valuation IS NOT INITIAL.
            LOOP AT mt_valuation ASSIGNING FIELD-SYMBOL(<ls_valuation>) USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
              IF line_exists( mt_valuation[ matnr = <ls_material>-matnr_target
                                            bwkey = <ls_valuation>-bwkey
                                            bwtar = <ls_valuation>-bwtar ] ).
                CONTINUE.
              ENDIF.
              ls_valuation-matnr = <ls_material>-matnr_target.
              ls_valuation-bwkey = <ls_valuation>-bwkey.
              ls_valuation-bwtar = <ls_valuation>-bwtar.
              INSERT ls_valuation INTO TABLE mt_valuation.
            ENDLOOP.
          ENDIF.

        WHEN sc_creation_types-excel.

          LOOP AT mt_werks ASSIGNING <ls_werks> USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
            IF line_exists( mt_werks[ matnr = <ls_material>-matnr_target
                                      werks = <ls_werks>-werks ] ).
              CONTINUE.
            ENDIF.

            ls_werks-matnr = <ls_material>-matnr_target.
            ls_werks-werks = <ls_werks>-werks.
            ls_werks-lgnum = <ls_werks>-lgnum.

            IF <ls_material>-minbe IS NOT INITIAL.
              ls_werks-minbe = <ls_material>-minbe.
            ELSE.
              ls_werks-minbe = <ls_werks>-minbe.
            ENDIF.

            IF <ls_material>-bstrf IS NOT INITIAL.
              ls_werks-bstrf = <ls_material>-bstrf.
            ELSE.
              ls_werks-bstrf = <ls_werks>-bstrf.
            ENDIF.

            IF <ls_material>-dismm IS NOT INITIAL.
              ls_werks-dismm = <ls_material>-dismm.
            ELSE.
              ls_werks-dismm = <ls_werks>-dismm.
            ENDIF.

            INSERT ls_werks INTO TABLE mt_werks.
            CLEAR ls_werks.

          ENDLOOP.

          LOOP AT mt_vkorg ASSIGNING <ls_vkorg> USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
            IF line_exists( mt_vkorg[ matnr = <ls_material>-matnr_target
                                      vkorg = <ls_vkorg>-vkorg
                                      vtweg = <ls_vkorg>-vtweg ] ).
              CONTINUE.
            ENDIF.

            ls_vkorg-matnr = <ls_material>-matnr_target.
            ls_vkorg-vkorg = <ls_vkorg>-vkorg.
            ls_vkorg-vtweg = <ls_vkorg>-vtweg.

            IF <ls_material>-kondm IS NOT INITIAL.
              ls_vkorg-kondm = <ls_material>-kondm.
            ELSE.
              ls_vkorg-kondm = <ls_vkorg>-kondm.
            ENDIF.

            IF <ls_material>-ktgrm IS NOT INITIAL.
              ls_vkorg-ktgrm = <ls_material>-ktgrm.
            ELSE.
              ls_vkorg-ktgrm = <ls_vkorg>-ktgrm.
            ENDIF.

            IF <ls_material>-mtpos IS NOT INITIAL.
              ls_vkorg-mtpos = <ls_material>-mtpos.
            ELSE.
              ls_vkorg-mtpos = <ls_vkorg>-mtpos.
            ENDIF.

            IF <ls_material>-tatyp IS NOT INITIAL.
              ls_vkorg-tatyp = <ls_material>-tatyp.
            ELSE.
              ls_vkorg-tatyp = <ls_vkorg>-tatyp.
            ENDIF.

            IF <ls_material>-taxkm IS NOT INITIAL.
              ls_vkorg-taxkm = <ls_material>-taxkm.
            ELSE.
              ls_vkorg-taxkm = <ls_vkorg>-taxkm.
            ENDIF.

            IF <ls_material>-taxim IS NOT INITIAL.
              ls_vkorg-taxim = <ls_material>-taxim.
            ELSE.
              ls_vkorg-taxim = <ls_vkorg>-taxim.
            ENDIF.

            INSERT ls_vkorg INTO TABLE mt_vkorg.
            CLEAR ls_vkorg.
          ENDLOOP.

          LOOP AT mt_valuation ASSIGNING <ls_valuation> USING KEY primary_key WHERE matnr EQ <ls_material>-matnr_source.
            IF line_exists( mt_valuation[ matnr = <ls_material>-matnr_target
                                          bwkey = <ls_valuation>-bwkey
                                          bwtar = <ls_valuation>-bwtar ] ).
              CONTINUE.
            ENDIF.

            ls_valuation-matnr = <ls_material>-matnr_target.
            ls_valuation-bwkey = <ls_valuation>-bwkey.
            ls_valuation-bwtar = <ls_valuation>-bwtar.

            IF <ls_material>-taxim IS NOT INITIAL.
              ls_valuation-stprs = <ls_material>-stprs.
            ELSE.
              ls_valuation-stprs = <ls_valuation>-stprs.
            ENDIF.

            INSERT ls_valuation INTO TABLE mt_valuation.
            CLEAR ls_valuation.
          ENDLOOP.

      ENDCASE.

    ENDLOOP.

  ENDMETHOD.

  METHOD assign_selected_werks.

    mo_alv_werks_source->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.

    DATA(lv_ask_taxim) = abap_false.
    " UB, 28.09.2026 10:50:25, Ç.No:127351
    " AKPA birleşme projesi
*    IF sy-mandt EQ '402'.
    LOOP AT mt_werks_source TRANSPORTING NO FIELDS
      WHERE werks+00(01) EQ '8'.
      EXIT.
    ENDLOOP.
    IF sy-subrc EQ 0.
      " UB, 28.09.2026 10:50:28, Ç.No:127351
      SELECT SINGLE matnr, aland, taxim
        FROM mlan
        WHERE matnr EQ @mr_material_selected->matnr_source
        INTO @DATA(ls_mlan).
      lv_ask_taxim = abap_true.
    ENDIF.

    READ TABLE lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>) INDEX 1.
    READ TABLE mt_werks_source ASSIGNING FIELD-SYMBOL(<ls_werks_source>) INDEX <ls_selected_row>-row_id.

    get_reference_material_data(
      EXPORTING
        iv_matnr               = mr_material_selected->matnr_source
        iv_plant               = <ls_werks_source>-werks
      IMPORTING
        es_plantdata           = DATA(ls_plantdata)
        es_warehousenumberdata = DATA(ls_warehousenumberdata) ).

    IF lines( lt_selected_rows ) GT 1.
      ls_plantdata-plant = '****'.
      ls_warehousenumberdata-whse_no = '***'.
    ENDIF.

    ask_werks_fields(
      EXPORTING
        iv_werks   = ls_plantdata-plant
        iv_minbe   = ls_plantdata-reorder_pt
        iv_bstrf   = ls_plantdata-round_val
        iv_dismm   = ls_plantdata-mrp_type
        iv_lgnum   = ls_warehousenumberdata-whse_no
      IMPORTING
        ev_aborted = DATA(lv_aborted)
        ev_minbe   = DATA(lv_minbe)
        ev_bstrf   = DATA(lv_bstrf)
        ev_dismm   = DATA(lv_dismm) ).

    IF lv_aborted IS NOT INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING <ls_selected_row>.
      READ TABLE mt_werks_source ASSIGNING <ls_werks_source> INDEX <ls_selected_row>-row_id.
      IF sy-subrc IS INITIAL.
        INSERT VALUE #( matnr = mr_material_selected->matnr_target
                        werks = <ls_werks_source>-werks
                        minbe = lv_minbe
                        bstrf = lv_bstrf
                        dismm = lv_dismm
                        lgnum = VALUE #( mt_lgnum[ werks = <ls_werks_source>-werks ]-lgnum OPTIONAL ) ) INTO TABLE mt_werks.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD unassign_selected_werks.
    mo_alv_werks_target->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>).
      READ TABLE mt_werks_target ASSIGNING FIELD-SYMBOL(<ls_werks_target>) INDEX <ls_selected_row>-row_id.
      IF sy-subrc IS INITIAL.
        DELETE mt_werks WHERE matnr EQ <ls_werks_target>-matnr
                          AND werks EQ <ls_werks_target>-werks.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD assign_selected_vkorg.

    mo_alv_vkorg_source->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.
    READ TABLE lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>) INDEX 1.
    READ TABLE mt_vkorg_source ASSIGNING FIELD-SYMBOL(<ls_vkorg_source>) INDEX <ls_selected_row>-row_id.

    DATA(lv_ask_taxim) = abap_false.
    " UB, 28.09.2026 10:50:25, Ç.No:127351
    " AKPA birleşme projesi
*    IF sy-mandt EQ '402'.
    LOOP AT mt_vkorg_source TRANSPORTING NO FIELDS
      WHERE vkorg+00(01) EQ '5'.
      EXIT.
    ENDLOOP.
    IF sy-subrc EQ 0.
      " UB, 28.09.2026 10:50:28, Ç.No:127351
      SELECT SINGLE matnr, aland, taxim
        FROM mlan
        WHERE matnr EQ @mr_material_selected->matnr_source
        INTO @DATA(ls_mlan).
      lv_ask_taxim = abap_true.
    ENDIF.

    IF lines( lt_selected_rows ) GT 1.
      <ls_vkorg_source>-vkorg = '****'.
      <ls_vkorg_source>-vtweg = '**'.
    ENDIF.

    get_reference_material_data(
      EXPORTING
        iv_matnr             = mr_material_selected->matnr_source
        iv_salesorg          = <ls_vkorg_source>-vkorg
        iv_distr_chan        = <ls_vkorg_source>-vtweg
      IMPORTING
        es_salesdata         = DATA(ls_salesdata)
        es_taxclassification = DATA(ls_taxclassification)
        ev_taxim             = DATA(lv_taxim_ask) ).

    ask_vkorg_fields(
      EXPORTING
        iv_vkorg     = ls_salesdata-sales_org
        iv_vtweg     = ls_salesdata-distr_chan
        iv_kondm     = ls_salesdata-mat_pr_grp
        iv_ktgrm     = ls_salesdata-acct_assgt
        iv_mtpos     = ls_salesdata-item_cat
        iv_tatyp     = ls_taxclassification-tax_type_1
        iv_taxkm     = ls_taxclassification-taxclass_1
        iv_taxim     = lv_taxim_ask
        iv_ask_taxim = lv_ask_taxim
      IMPORTING
        ev_aborted   = DATA(lv_aborted)
        ev_kondm     = DATA(lv_kondm)
        ev_ktgrm     = DATA(lv_ktgrm)
        ev_mtpos     = DATA(lv_mtpos)
        ev_tatyp     = DATA(lv_tatyp)
        ev_taxkm     = DATA(lv_taxkm)
        ev_taxim     = DATA(lv_taxim) ).

    IF lv_aborted IS NOT INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING <ls_selected_row>.

      READ TABLE mt_vkorg_source ASSIGNING <ls_vkorg_source> INDEX <ls_selected_row>-row_id.
      IF sy-subrc IS INITIAL.
        INSERT VALUE #( matnr = mr_material_selected->matnr_target
                        vkorg = <ls_vkorg_source>-vkorg
                        vtweg = <ls_vkorg_source>-vtweg
                        kondm = lv_kondm
                        ktgrm = lv_ktgrm
                        mtpos = lv_mtpos
                        tatyp = lv_tatyp
                        taxkm = lv_taxkm
                        taxim = lv_taxim ) INTO TABLE mt_vkorg.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD unassign_selected_vkorg.
    mo_alv_vkorg_target->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>).
      READ TABLE mt_vkorg_target ASSIGNING FIELD-SYMBOL(<ls_vkorg_target>) INDEX <ls_selected_row>-row_id.
      DELETE mt_vkorg WHERE matnr EQ <ls_vkorg_target>-matnr
                        AND vkorg EQ <ls_vkorg_target>-vkorg
                        AND vtweg EQ <ls_vkorg_target>-vtweg.
    ENDLOOP.
  ENDMETHOD.

  METHOD assign_selected_valuation.

    mo_alv_valuation_source->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.

    READ TABLE lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>) INDEX 1.
    READ TABLE mt_valuation_source ASSIGNING FIELD-SYMBOL(<ls_valuation_source>) INDEX <ls_selected_row>-row_id.

    DATA(lv_ask_taxim) = abap_false.
    " UB, 28.09.2026 10:50:25, Ç.No:127351
    " AKPA birleşme projesi
*    IF sy-mandt EQ '402'.
    LOOP AT mt_valuation_source TRANSPORTING NO FIELDS
      WHERE bwkey+00(01) EQ '8'.
      EXIT.
    ENDLOOP.
    IF sy-subrc EQ 0.
      " UB, 28.09.2026 10:50:28, Ç.No:127351
      SELECT SINGLE matnr, aland, taxim
        FROM mlan
        WHERE matnr EQ @mr_material_selected->matnr_source
        INTO @DATA(ls_mlan).
      lv_ask_taxim = abap_true.
    ENDIF.

    get_reference_material_data(
      EXPORTING
        iv_matnr         = mr_material_selected->matnr_source
        iv_val_area      = <ls_valuation_source>-bwkey
        iv_val_type      = <ls_valuation_source>-bwtar
      IMPORTING
        es_valuationdata = DATA(ls_valuationdata) ).

    IF lines( lt_selected_rows ) GT 1.
      ls_valuationdata-val_area = '****'.
      ls_valuationdata-val_type = '**********'.
    ENDIF.

    ask_valuation_fields(
      EXPORTING
        iv_bwkey   = ls_valuationdata-val_area
        iv_bwtar   = ls_valuationdata-val_type
        iv_stprs   = CONV #( ls_valuationdata-std_price )
      IMPORTING
        ev_aborted = DATA(lv_aborted)
        ev_stprs   = DATA(lv_stprs) ).

    IF lv_aborted IS NOT INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING <ls_selected_row>.

      READ TABLE mt_valuation_source ASSIGNING <ls_valuation_source> INDEX <ls_selected_row>-row_id.
      IF sy-subrc IS INITIAL.
        INSERT VALUE #( matnr = mr_material_selected->matnr_target
                        bwkey = <ls_valuation_source>-bwkey
                        bwtar = <ls_valuation_source>-bwtar
                        stprs = lv_stprs ) INTO TABLE mt_valuation.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD unassign_selected_valuation.
    mo_alv_valuation_target->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    IF lt_selected_rows IS INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>).
      READ TABLE mt_valuation_target ASSIGNING FIELD-SYMBOL(<ls_valuation_target>) INDEX <ls_selected_row>-row_id.
      DELETE mt_valuation WHERE matnr EQ <ls_valuation_target>-matnr
                            AND bwkey EQ <ls_valuation_target>-bwkey
                            AND bwtar EQ <ls_valuation_target>-bwtar.
    ENDLOOP.
  ENDMETHOD.

  METHOD create_and_extend_material.

    DATA: lt_bapi_headdata             TYPE TABLE OF bapie1matheader,

          lt_bapi_clientdata           TYPE TABLE OF bapie1mara,
          lt_bapi_clientdatax          TYPE TABLE OF bapie1marax,

          lt_bapi_plantdata            TYPE TABLE OF bapie1marc,
          lt_bapi_plantdatax           TYPE TABLE OF bapie1marcx,

          lt_bapi_warehousenumberdata  TYPE TABLE OF bapie1mlgn,
          lt_bapi_warehousenumberdatax TYPE TABLE OF bapie1mlgnx,

          lt_bapi_valuationdata        TYPE TABLE OF bapie1mbew,
          lt_bapi_valuationdatax       TYPE TABLE OF bapie1mbewx,

          lt_bapi_salesdata            TYPE TABLE OF bapie1mvke,
          lt_bapi_salesdatax           TYPE TABLE OF bapie1mvkex,

          lt_bapi_storagelocationdata  TYPE TABLE OF bapie1mard,
          lt_bapi_storagelocationdatax TYPE TABLE OF bapie1mardx,

          lt_bapi_taxclassifications   TYPE TABLE OF bapie1mlan,

          lt_bapi_materialdescription  TYPE TABLE OF bapie1makt,

          lt_bapi_unitsofmeasure       TYPE TABLE OF bapie1marm,
          lt_bapi_unitsofmeasurex      TYPE TABLE OF bapie1marmx,

          lt_bapi_extensionin          TYPE TABLE OF bapie1parex,
          lt_bapi_extensioninx         TYPE TABLE OF bapie1parexx.

    DATA: ls_client_data         TYPE bapi_mara_ga,
          lt_materialdescription TYPE TABLE OF bapi_makt_ga,
          lt_unitsofmeasure      TYPE TABLE OF bapi_marm_ga,

          function_insert        TYPE string VALUE 'INS',
          function_update        TYPE string VALUE 'UPD'.

    READ TABLE mt_material ASSIGNING FIELD-SYMBOL(<ls_material>) WITH KEY matnr_source = iv_matnr_source matnr_target = iv_matnr.

    CLEAR: ls_client_data,
           lt_materialdescription,
           lt_unitsofmeasure,
           et_return.

    CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        material_long       = <ls_material>-matnr_source
      IMPORTING
        clientdata          = ls_client_data
      TABLES
        materialdescription = lt_materialdescription
        unitsofmeasure      = lt_unitsofmeasure
        return              = et_return.

    LOOP AT et_return TRANSPORTING NO FIELDS WHERE type CA 'EAX'.
      EXIT.
    ENDLOOP.
    IF sy-subrc IS INITIAL.
      RETURN.
    ENDIF.

    DATA(lv_matnr_target) = <ls_material>-matnr_target.
    IF <ls_material>-matnr_target CP '$*'.
      lv_matnr_target = material_get_number_next( ls_client_data-matl_type ).
    ENDIF.

    CLEAR lt_bapi_headdata.
    APPEND VALUE #( ) TO lt_bapi_headdata ASSIGNING FIELD-SYMBOL(<ls_bapi_headdata>).
    <ls_bapi_headdata> = CORRESPONDING #( ls_client_data ).
    <ls_bapi_headdata>-function      = function_insert.
    <ls_bapi_headdata>-material      = lv_matnr_target.
    <ls_bapi_headdata>-material_long = lv_matnr_target.
    <ls_bapi_headdata>-basic_view    = abap_true.
    <ls_bapi_headdata>-purchase_view = abap_true.
    <ls_bapi_headdata>-sales_view    = abap_true.
    <ls_bapi_headdata>-account_view  = abap_true.
    <ls_bapi_headdata>-mrp_view      = abap_true.
    <ls_bapi_headdata>-prt_view      = abap_true.
    <ls_bapi_headdata>-storage_view  = abap_true.
    <ls_bapi_headdata>-cost_view     = abap_true.
    <ls_bapi_headdata>-quality_view  = abap_true.

    CLEAR lt_bapi_clientdata.
    APPEND CORRESPONDING #( ls_client_data ) TO lt_bapi_clientdata ASSIGNING FIELD-SYMBOL(<ls_bapi_clientdata>).
    <ls_bapi_clientdata>-function      = function_insert.
    <ls_bapi_clientdata>-material      = lv_matnr_target.
    <ls_bapi_clientdata>-material_long = lv_matnr_target.

    CLEAR lt_bapi_clientdatax.
    APPEND VALUE #( ) TO lt_bapi_clientdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_clientdatax>).
    <ls_bapi_clientdatax>-function      = function_insert.
    <ls_bapi_clientdatax>-material      = lv_matnr_target.
    <ls_bapi_clientdatax>-material_long = lv_matnr_target.

    fill_bapix( REF #( <ls_bapi_clientdatax> ) ).
    CLEAR: <ls_bapi_clientdatax>-batch_mgmt.

    CLEAR: lt_bapi_plantdata,
           lt_bapi_plantdatax.
    LOOP AT mt_werks ASSIGNING FIELD-SYMBOL(<ls_werks_target>) WHERE matnr EQ <ls_material>-matnr_target.

      DATA(ls_plantdata) = VALUE bapi_marc_ga( ).
      DATA(ls_warehousenumberdata) = VALUE bapi_mlgn_ga( ).
      CALL FUNCTION 'BAPI_MATERIAL_GET_ALL'
        EXPORTING
          material_long       = <ls_material>-matnr_source
          plant               = <ls_werks_target>-werks
          whsenumber          = <ls_werks_target>-lgnum
        IMPORTING
          plantdata           = ls_plantdata
          warehousenumberdata = ls_warehousenumberdata.

      IF ls_plantdata IS NOT INITIAL.

        DATA(ls_plantdata_target) = VALUE bapi_marc_ga( ).
        DATA(ls_warehousenumberdata_target) = VALUE bapi_mlgn_ga( ).
        CALL FUNCTION 'BAPI_MATERIAL_GET_ALL'
          EXPORTING
            material_long       = <ls_material>-matnr_target
            plant               = <ls_werks_target>-werks
            whsenumber          = <ls_werks_target>-lgnum
          IMPORTING
            plantdata           = ls_plantdata_target
            warehousenumberdata = ls_warehousenumberdata_target.

        APPEND INITIAL LINE TO lt_bapi_plantdata ASSIGNING FIELD-SYMBOL(<ls_bapi_plantdata>).
        APPEND INITIAL LINE TO lt_bapi_plantdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_plantdatax>).

        <ls_bapi_plantdata> = CORRESPONDING #( ls_plantdata ).

        IF ls_plantdata_target IS NOT INITIAL.
          <ls_bapi_plantdata>-function = function_update.
        ELSE.
          <ls_bapi_plantdata>-function = function_insert.
        ENDIF.

        <ls_bapi_plantdata>-material      = lv_matnr_target.
        <ls_bapi_plantdata>-material_long = lv_matnr_target.
        <ls_bapi_plantdata>-plant         = ls_plantdata-plant.

        <ls_bapi_plantdatax>-function      = <ls_bapi_plantdata>-function.
        <ls_bapi_plantdatax>-material      = lv_matnr_target.
        <ls_bapi_plantdatax>-material_long = lv_matnr_target.
        <ls_bapi_plantdatax>-plant         = ls_plantdata-plant.

        fill_bapix( REF #( <ls_bapi_plantdatax> ) ).

        CLEAR: <ls_bapi_plantdatax>-del_flag,
               <ls_bapi_plantdatax>-reorder_pt,
               <ls_bapi_plantdatax>-round_val.

        IF <ls_werks_target>-minbe IS NOT INITIAL.
          <ls_bapi_plantdata>-reorder_pt  = <ls_werks_target>-minbe.
          <ls_bapi_plantdatax>-reorder_pt = abap_true.
        ENDIF.

        IF <ls_werks_target>-bstrf IS NOT INITIAL.
          <ls_bapi_plantdata>-round_val  = <ls_werks_target>-bstrf.
          <ls_bapi_plantdatax>-round_val = abap_true.
        ENDIF.

        IF <ls_werks_target>-dismm IS NOT INITIAL.
          <ls_bapi_plantdata>-mrp_type  = <ls_werks_target>-dismm.
          <ls_bapi_plantdatax>-mrp_type = abap_true.
        ENDIF.

      ENDIF.

      IF ls_warehousenumberdata-whse_no IS NOT INITIAL.

        APPEND INITIAL LINE TO lt_bapi_warehousenumberdata ASSIGNING FIELD-SYMBOL(<ls_bapi_warehousenumberdata>).
        APPEND INITIAL LINE TO lt_bapi_warehousenumberdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_warehousenumberdatax>).

        <ls_bapi_warehousenumberdata> = CORRESPONDING #( ls_warehousenumberdata ).

        <ls_bapi_warehousenumberdata>-function      = function_insert.
        <ls_bapi_warehousenumberdata>-material      = lv_matnr_target.
        <ls_bapi_warehousenumberdata>-material_long = lv_matnr_target.

        <ls_bapi_warehousenumberdatax>-function      = function_insert.
        <ls_bapi_warehousenumberdatax>-material      = lv_matnr_target.
        <ls_bapi_warehousenumberdatax>-material_long = lv_matnr_target.

        <ls_bapi_warehousenumberdatax>-whse_no = ls_warehousenumberdata-whse_no.

        fill_bapix( REF #( <ls_bapi_warehousenumberdatax> ) ).

        CLEAR: <ls_bapi_warehousenumberdatax>-del_flag.

      ENDIF.

    ENDLOOP.

    CLEAR: lt_bapi_salesdata,
           lt_bapi_salesdatax,
           lt_bapi_taxclassifications.
    LOOP AT mt_vkorg ASSIGNING FIELD-SYMBOL(<ls_vkorg_target>) WHERE matnr EQ <ls_material>-matnr_target.

      DATA(ls_salesdata) = VALUE bapi_mvke_ga( ).
      DATA(lt_taxclassifications) = VALUE txs_mat_tax_classif_tty( ).
      CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
        EXPORTING
          material_long      = <ls_material>-matnr_source
          salesorg           = <ls_vkorg_target>-vkorg
          distr_chan         = <ls_vkorg_target>-vtweg
        IMPORTING
          salesdata          = ls_salesdata
        TABLES
          taxclassifications = lt_taxclassifications.

      IF ls_salesdata IS NOT INITIAL.

        DATA(ls_salesdata_target) = VALUE bapi_mvke_ga( ).
        DATA(lt_taxclassifications_target) = VALUE txs_mat_tax_classif_tty( ).
        CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
          EXPORTING
            material_long      = <ls_material>-matnr_target
            salesorg           = <ls_vkorg_target>-vkorg
            distr_chan         = <ls_vkorg_target>-vtweg
          IMPORTING
            salesdata          = ls_salesdata_target
          TABLES
            taxclassifications = lt_taxclassifications_target.

        APPEND INITIAL LINE TO lt_bapi_salesdata ASSIGNING FIELD-SYMBOL(<ls_bapi_salesdata>).
        APPEND INITIAL LINE TO lt_bapi_salesdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_salesdatax>).

        <ls_bapi_salesdata> = CORRESPONDING #( ls_salesdata ).

        IF ls_salesdata_target IS NOT INITIAL.
          <ls_bapi_salesdata>-function = function_update.
        ELSE.
          <ls_bapi_salesdata>-function = function_insert.
        ENDIF.

        <ls_bapi_salesdata>-material      = lv_matnr_target.
        <ls_bapi_salesdata>-material_long = lv_matnr_target.
        <ls_bapi_salesdata>-sales_org     = ls_salesdata-sales_org.
        <ls_bapi_salesdata>-distr_chan    = ls_salesdata-distr_chan.

        <ls_bapi_salesdatax>-function      = <ls_bapi_salesdata>-function.
        <ls_bapi_salesdatax>-material      = lv_matnr_target.
        <ls_bapi_salesdatax>-material_long = lv_matnr_target.
        <ls_bapi_salesdatax>-sales_org     = ls_salesdata-sales_org.
        <ls_bapi_salesdatax>-distr_chan    = ls_salesdata-distr_chan.
        <ls_bapi_salesdatax>-del_flag      = abap_true.

        fill_bapix( REF #( <ls_bapi_salesdatax> ) ).

        CLEAR: <ls_bapi_salesdatax>-del_flag,
               <ls_bapi_salesdatax>-mat_pr_grp,
               <ls_bapi_salesdatax>-acct_assgt,
               <ls_bapi_salesdatax>-item_cat.

        IF <ls_vkorg_target>-kondm IS NOT INITIAL.
          <ls_bapi_salesdata>-mat_pr_grp  = <ls_vkorg_target>-kondm.
          <ls_bapi_salesdatax>-mat_pr_grp = abap_true.
        ENDIF.

        IF <ls_vkorg_target>-ktgrm IS NOT INITIAL.
          <ls_bapi_salesdata>-acct_assgt  = <ls_vkorg_target>-ktgrm.
          <ls_bapi_salesdatax>-acct_assgt = abap_true.
        ENDIF.

        IF <ls_vkorg_target>-mtpos IS NOT INITIAL.
          <ls_bapi_salesdata>-item_cat  = <ls_vkorg_target>-mtpos.
          <ls_bapi_salesdatax>-item_cat = abap_true.
        ENDIF.

      ENDIF.

      LOOP AT lt_taxclassifications ASSIGNING FIELD-SYMBOL(<ls_taxclassification>).

        APPEND INITIAL LINE TO lt_bapi_taxclassifications ASSIGNING FIELD-SYMBOL(<ls_bapi_taxclassification>).

        <ls_bapi_taxclassification> = CORRESPONDING #( <ls_taxclassification> ).

        <ls_bapi_taxclassification>-function      = function_insert.
        <ls_bapi_taxclassification>-material      = lv_matnr_target.
        <ls_bapi_taxclassification>-material_long = lv_matnr_target.

        IF <ls_vkorg_target>-tatyp IS NOT INITIAL.
          <ls_bapi_taxclassification>-tax_type_1 = <ls_vkorg_target>-tatyp.
        ENDIF.

        IF <ls_vkorg_target>-taxkm IS NOT INITIAL.
          <ls_bapi_taxclassification>-taxclass_1 = <ls_vkorg_target>-taxkm.
        ENDIF.

        IF <ls_vkorg_target>-taxim IS NOT INITIAL.
          <ls_bapi_taxclassification>-tax_ind = <ls_vkorg_target>-taxim.
        ENDIF.

      ENDLOOP.

    ENDLOOP.

    SORT lt_bapi_taxclassifications BY function material depcountry depcountry_iso.
    DELETE ADJACENT DUPLICATES FROM lt_bapi_taxclassifications COMPARING material depcountry depcountry_iso.

    SELECT mbew~*
      FROM @mt_valuation AS valuation
      JOIN mbew ON mbew~matnr EQ valuation~matnr
               AND mbew~bwkey EQ valuation~bwkey
               AND mbew~bwtar EQ valuation~bwtar
     ORDER BY mbew~matnr,
              mbew~bwkey,
              mbew~bwtar
      INTO TABLE @DATA(lt_mbew_db).

    CLEAR: lt_bapi_valuationdata,
           lt_bapi_valuationdatax.
    LOOP AT mt_valuation ASSIGNING FIELD-SYMBOL(<ls_valuation_target>) WHERE matnr EQ <ls_material>-matnr_target.

      DATA(ls_valuationdata) = VALUE bapi_mbew_ga( ).
      CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
        EXPORTING
          material_long = <ls_material>-matnr_source
          val_area      = <ls_valuation_target>-bwkey
          val_type      = <ls_valuation_target>-bwtar
        IMPORTING
          valuationdata = ls_valuationdata.

      IF ls_valuationdata IS NOT INITIAL.

        DATA(ls_valuationdata_target) = VALUE bapi_mbew_ga( ).
        CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
          EXPORTING
            material_long = <ls_material>-matnr_target
            val_area      = <ls_valuation_target>-bwkey
            val_type      = <ls_valuation_target>-bwtar
          IMPORTING
            valuationdata = ls_valuationdata_target.

        APPEND INITIAL LINE TO lt_bapi_valuationdata ASSIGNING FIELD-SYMBOL(<ls_bapi_valuationdata>).
        APPEND INITIAL LINE TO lt_bapi_valuationdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_valuationdatax>).

        <ls_bapi_valuationdata> = CORRESPONDING #( ls_valuationdata ).

        IF ls_valuationdata_target IS NOT INITIAL.
          <ls_bapi_valuationdata>-function = function_update.
        ELSE.
          <ls_bapi_valuationdata>-function = function_insert.
        ENDIF.

        <ls_bapi_valuationdata>-material      = lv_matnr_target.
        <ls_bapi_valuationdata>-material_long = lv_matnr_target.

        READ TABLE lt_mbew_db ASSIGNING FIELD-SYMBOL(<ls_mbew_db>) WITH KEY matnr = <ls_material>-matnr_source
                                                                            bwkey = ls_valuationdata-val_area
                                                                            bwtar = ls_valuationdata-val_type BINARY SEARCH.

        IF sy-subrc IS INITIAL.
          <ls_bapi_valuationdata>-price_ctrl = <ls_mbew_db>-vprsv.
          <ls_bapi_valuationdata>-pr_ctrl_py = <ls_mbew_db>-vjvpr.
          <ls_bapi_valuationdata>-pr_ctrl_pp = <ls_mbew_db>-vmvpr.
        ENDIF.

        <ls_bapi_valuationdatax>-function      = <ls_bapi_valuationdata>-function.
        <ls_bapi_valuationdatax>-material      = lv_matnr_target.
        <ls_bapi_valuationdatax>-material_long = lv_matnr_target.
        <ls_bapi_valuationdatax>-val_area      = ls_valuationdata-val_area.
        <ls_bapi_valuationdatax>-val_type      = ls_valuationdata-val_type.

        fill_bapix( REF #( <ls_bapi_valuationdatax> ) ).

        CLEAR: <ls_bapi_valuationdatax>-del_flag,
               <ls_bapi_valuationdatax>-moving_pr,
               <ls_bapi_valuationdatax>-std_pr_pp,
               <ls_bapi_valuationdatax>-std_price,
               <ls_bapi_valuationdatax>-mov_pr_py,
               <ls_bapi_valuationdatax>-std_pr_py,
               <ls_bapi_valuationdatax>-mov_pr_pp,
               <ls_bapi_valuationdatax>-std_price.

        IF ls_valuationdata-price_ctrl = 'S' .
          <ls_bapi_valuationdata>-std_price  = <ls_valuation_target>-stprs.
          <ls_bapi_valuationdatax>-std_price = abap_true.
        ENDIF.

      ENDIF.

    ENDLOOP.

    CLEAR: lt_bapi_storagelocationdata,
           lt_bapi_storagelocationdatax.
    LOOP AT mt_lgort ASSIGNING FIELD-SYMBOL(<ls_lgort>) WHERE matnr EQ <ls_material>-matnr_source.

      DATA(ls_storagelocationdata) = VALUE bapi_mard_ga( ).
      CALL FUNCTION 'BAPI_MATERIAL_GET_ALL'
        EXPORTING
          material_long       = <ls_material>-matnr_source
          plant               = <ls_lgort>-werks
          stge_loc            = <ls_lgort>-lgort
        IMPORTING
          storagelocationdata = ls_storagelocationdata.

      IF ls_storagelocationdata IS NOT INITIAL.

        APPEND INITIAL LINE TO lt_bapi_storagelocationdata ASSIGNING FIELD-SYMBOL(<ls_bapi_storagelocationdata>).
        APPEND INITIAL LINE TO lt_bapi_storagelocationdatax ASSIGNING FIELD-SYMBOL(<ls_bapi_storagelocationdatax>).

        <ls_bapi_storagelocationdata> = CORRESPONDING #( ls_storagelocationdata ).

        <ls_bapi_storagelocationdata>-function      = function_insert.
        <ls_bapi_storagelocationdata>-material      = lv_matnr_target.
        <ls_bapi_storagelocationdata>-material_long = lv_matnr_target.
        <ls_bapi_storagelocationdata>-plant         = ls_storagelocationdata-plant.
        <ls_bapi_storagelocationdata>-stge_loc      = ls_storagelocationdata-stge_loc.

        <ls_bapi_storagelocationdatax>-function      = function_insert.
        <ls_bapi_storagelocationdatax>-material      = lv_matnr_target.
        <ls_bapi_storagelocationdatax>-material_long = lv_matnr_target.
        <ls_bapi_storagelocationdatax>-plant         = ls_storagelocationdata-plant.
        <ls_bapi_storagelocationdatax>-stge_loc      = ls_storagelocationdata-stge_loc.

        fill_bapix( REF #( <ls_bapi_storagelocationdatax> ) ).
        CLEAR: <ls_bapi_storagelocationdatax>-del_flag.

      ENDIF.

    ENDLOOP.

    CLEAR: lt_bapi_unitsofmeasure,
           lt_bapi_unitsofmeasurex.
    LOOP AT lt_unitsofmeasure ASSIGNING FIELD-SYMBOL(<ls_unitsofmeasure>).

      APPEND INITIAL LINE TO lt_bapi_unitsofmeasure ASSIGNING FIELD-SYMBOL(<ls_bapi_unitsofmeasure>).
      APPEND INITIAL LINE TO lt_bapi_unitsofmeasurex ASSIGNING FIELD-SYMBOL(<ls_bapi_unitsofmeasurex>).

      <ls_bapi_unitsofmeasure> = CORRESPONDING #( <ls_unitsofmeasure> ).

      <ls_bapi_unitsofmeasure>-function      = function_insert.
      <ls_bapi_unitsofmeasure>-material      = lv_matnr_target.
      <ls_bapi_unitsofmeasure>-material_long = lv_matnr_target.
      <ls_bapi_unitsofmeasure>-alt_unit      = <ls_unitsofmeasure>-alt_unit.
      <ls_bapi_unitsofmeasure>-alt_unit_iso  = <ls_unitsofmeasure>-alt_unit_iso.

      <ls_bapi_unitsofmeasurex>-function       = function_insert.
      <ls_bapi_unitsofmeasurex>-material       = lv_matnr_target.
      <ls_bapi_unitsofmeasurex>-material_long  = lv_matnr_target.
      <ls_bapi_unitsofmeasurex>-alt_unit       = <ls_unitsofmeasure>-alt_unit.
      <ls_bapi_unitsofmeasurex>-alt_unit_iso   = <ls_unitsofmeasure>-alt_unit_iso.

      <ls_bapi_unitsofmeasurex>-gross_wt       = abap_true.
      <ls_bapi_unitsofmeasurex>-unit_of_wt     = abap_true.
      <ls_bapi_unitsofmeasurex>-unit_of_wt_iso = abap_true.

    ENDLOOP.

    CLEAR lt_bapi_materialdescription.
    LOOP AT lt_materialdescription ASSIGNING FIELD-SYMBOL(<ls_materialdescription>).
      APPEND VALUE #( function      = function_insert
                      material      = lv_matnr_target
                      material_long = lv_matnr_target
                      langu         = <ls_materialdescription>-langu
                      langu_iso     = <ls_materialdescription>-langu_iso
                      matl_desc     = <ls_material>-maktx_target ) TO lt_bapi_materialdescription.
    ENDLOOP.

    SELECT SINGLE zzurunh
      FROM mara
     WHERE matnr EQ @<ls_material>-matnr_source
     INTO @DATA(lv_zzurunh).

    zcl_im_badi_material_check=>set_extension_data( VALUE #( ( material = lv_matnr_target
                                                               zzurunh  = lv_zzurunh
                                                               zziy     = <ls_material>-zziy ) ) ).

*    lt_bapi_extensionin = VALUE #( ( function   = function_insert
*                                     material   = lv_matnr_target
*                                     structure  = 'BAPI_TE_MARA'
*                                     valuepart1 = VALUE bapi_te_mara( material = lv_matnr_target
*                                                                      zzurunh  = lv_zzurunh
*                                                                      zziy     = <ls_material>-zziy ) ) ).
*
*    lt_bapi_extensioninx = VALUE #( ( function   = function_insert
*                                      material   = lv_matnr_target
*                                      structure  = 'BAPI_TE_MARAX'
*                                      valuepart1 = VALUE bapi_te_mara( material = lv_matnr_target
*                                                                       zzurunh  = COND #( WHEN lv_zzurunh IS NOT INITIAL
*                                                                                          THEN abap_true
*                                                                                          ELSE abap_false )
*                                                                       zziy     = abap_true ) ) ).

    DATA(ls_return) = VALUE bapiret2( ).

    sy-uname = 'JOBMM'.
    CALL FUNCTION 'BAPI_MATERIAL_SAVEREPLICA' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        noappllog            = abap_false
        nochangedoc          = abap_false
        testrun              = abap_false
        inpfldcheck          = abap_false
      IMPORTING
        return               = ls_return
      TABLES
        headdata             = lt_bapi_headdata
        clientdata           = lt_bapi_clientdata
        clientdatax          = lt_bapi_clientdatax
        plantdata            = lt_bapi_plantdata
        plantdatax           = lt_bapi_plantdatax
        warehousenumberdata  = lt_bapi_warehousenumberdata
        warehousenumberdatax = lt_bapi_warehousenumberdatax
        valuationdata        = lt_bapi_valuationdata
        valuationdatax       = lt_bapi_valuationdatax
        salesdata            = lt_bapi_salesdata
        salesdatax           = lt_bapi_salesdatax
        storagelocationdata  = lt_bapi_storagelocationdata
        storagelocationdatax = lt_bapi_storagelocationdatax
        taxclassifications   = lt_bapi_taxclassifications
        materialdescription  = lt_bapi_materialdescription
        unitsofmeasure       = lt_bapi_unitsofmeasure
        unitsofmeasurex      = lt_bapi_unitsofmeasurex
        extensionin          = lt_bapi_extensionin
        extensioninx         = lt_bapi_extensioninx
        returnmessages       = et_return.

    zcl_im_badi_material_check=>free_extension_data( ).

    IF ls_return-type CA 'EAX'.
      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      RETURN.
    ENDIF.

    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        wait = abap_true.

    LOOP AT mt_werks INTO DATA(ls_werks) WHERE matnr EQ <ls_material>-matnr_target.
      ls_werks-matnr = lv_matnr_target.
      INSERT ls_werks INTO TABLE mt_werks.
    ENDLOOP.
    DELETE mt_werks WHERE matnr EQ <ls_material>-matnr_target.

    LOOP AT mt_vkorg INTO DATA(ls_vkorg) WHERE matnr EQ <ls_material>-matnr_target.
      ls_vkorg-matnr = lv_matnr_target.
      INSERT ls_vkorg INTO TABLE mt_vkorg.
    ENDLOOP.
    DELETE mt_vkorg WHERE matnr EQ <ls_material>-matnr_target.

    LOOP AT mt_valuation INTO DATA(ls_valuation) WHERE matnr EQ <ls_material>-matnr_target.
      ls_valuation-matnr = lv_matnr_target.
      INSERT ls_valuation INTO TABLE mt_valuation.
    ENDLOOP.
    DELETE mt_valuation WHERE matnr EQ <ls_material>-matnr_target.

    <ls_material>-matnr_target = lv_matnr_target.

  ENDMETHOD.

  METHOD material_get_number_next.
    CALL FUNCTION 'MATERIAL_NUMBER_GET_NEXT'
      EXPORTING
        materialart          = iv_mtart
      IMPORTING
        materialnr           = rv_matnr
      EXCEPTIONS
        no_internal_interval = 1
        type_not_found       = 2
        OTHERS               = 3.
  ENDMETHOD.


  METHOD create.

    mo_alv_material->get_selected_rows(
      IMPORTING
        et_row_no = DATA(lt_selected_rows) ).

    LOOP AT lt_selected_rows ASSIGNING FIELD-SYMBOL(<ls_selected_row>).
      READ TABLE mt_material ASSIGNING FIELD-SYMBOL(<ls_material>) INDEX <ls_selected_row>-row_id.
      IF sy-subrc IS INITIAL.
        create_and_extend_material(
          EXPORTING
            iv_matnr_source = <ls_material>-matnr_source
            iv_matnr        = <ls_material>-matnr_target
          IMPORTING
            et_return       = <ls_material>-messages ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD fill_bapix.

    DATA lo_strucdescr TYPE REF TO cl_abap_structdescr.

    FIELD-SYMBOLS: <fs_line> TYPE any.

    lo_strucdescr ?= cl_abap_typedescr=>describe_by_data_ref( ir_data ).

    ASSIGN ir_data->* TO <fs_line>.

    LOOP AT lo_strucdescr->get_components( ) ASSIGNING FIELD-SYMBOL(<ls_component>).

      IF <ls_component>-type->absolute_name EQ '\TYPE=BAPIUPDATE'.
      ELSE.
        CONTINUE.
      ENDIF.

      ASSIGN COMPONENT <ls_component>-name OF STRUCTURE <fs_line> TO FIELD-SYMBOL(<lv_value>).
      IF sy-subrc IS INITIAL.
        <lv_value> = abap_true.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.

  METHOD popup_get_values.
    CALL FUNCTION 'POPUP_GET_VALUES'
      EXPORTING
        no_value_check  = iv_no_value_check
        popup_title     = iv_popup_title
        start_column    = iv_start_column
        start_row       = iv_start_row
      IMPORTING
        returncode      = ev_returncode
      TABLES
        fields          = ct_fields
      EXCEPTIONS
        error_in_fields = 1
        OTHERS          = 2.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDMETHOD.


  METHOD ask_werks_fields.

    DATA(lt_fields) = VALUE ty_sval( ).

    APPEND VALUE #( tabname   = 'MARA'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_source
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_target
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'WERKS'
                    field_obl = abap_false
                    value     = iv_werks
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'MINBE'
                    field_obl = abap_true
                    value     = iv_minbe ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'BSTRF'
                    field_obl = abap_true
                    value     = iv_bstrf ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'DISMM'
                    field_obl = abap_true
                    value     = iv_dismm ) TO lt_fields.

    APPEND VALUE #( tabname    = 'MLGN'
                    fieldname  = 'LGNUM'
                    field_obl  = abap_false
                    value      = iv_lgnum
                    field_attr = '02'  ) TO lt_fields.

    popup_get_values(
      EXPORTING
        iv_no_value_check = abap_false
        iv_popup_title    = 'Aşağıdaki bilgileri doldurunuz'
      IMPORTING
        ev_returncode     = DATA(lv_return_code)
      CHANGING
        ct_fields         = lt_fields ).

    ev_aborted = xsdbool( lv_return_code IS NOT INITIAL ).

    IF ev_aborted EQ abap_true.
      RETURN.
    ENDIF.

    ev_minbe = VALUE #( lt_fields[ tabname   = 'MARC'
                                   fieldname = 'MINBE' ]-value ).

    ev_bstrf = VALUE #( lt_fields[ tabname   = 'MARC'
                                   fieldname = 'BSTRF' ]-value ).

    ev_dismm = VALUE #( lt_fields[ tabname   = 'MARC'
                                   fieldname = 'DISMM' ]-value ).

  ENDMETHOD.


  METHOD ask_vkorg_fields.

    DATA(lt_fields) = VALUE ty_sval( ).

    APPEND VALUE #( tabname   = 'MARA'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_source
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_target
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MVKE'
                    fieldname = 'VKORG'
                    field_obl = abap_false
                    value     = iv_vkorg
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MVKE'
                    fieldname = 'VTWEG'
                    field_obl = abap_false
                    value     = iv_vtweg
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MVKE'
                    fieldname = 'KONDM'
                    field_obl = abap_true
                    value     = iv_kondm ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MVKE'
                    fieldname = 'KTGRM'
                    field_obl = abap_true
                    value     = iv_ktgrm ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MVKE'
                    fieldname = 'MTPOS'
                    field_obl = abap_true
                    value     = iv_mtpos ) TO lt_fields.

    APPEND VALUE #( tabname   = 'TSKM'
                    fieldname = 'TATYP'
                    field_obl = abap_true
                    value     = iv_tatyp ) TO lt_fields.

    APPEND VALUE #( tabname   = 'TSKM'
                    fieldname = 'TAXKM'
                    field_obl = abap_true
                    value     = iv_taxkm ) TO lt_fields.

    IF iv_ask_taxim EQ abap_true.
      APPEND VALUE #( tabname   = 'MLAN'
                      fieldname = 'TAXIM'
                      field_obl = abap_true
                      value     = iv_taxim ) TO lt_fields.
    ENDIF.

    popup_get_values(
      EXPORTING
        iv_no_value_check = abap_false
        iv_popup_title    = 'Aşağıdaki bilgileri doldurunuz'
      IMPORTING
        ev_returncode     = DATA(lv_return_code)
      CHANGING
        ct_fields         = lt_fields ).

    ev_aborted = xsdbool( lv_return_code IS NOT INITIAL ).

    IF ev_aborted EQ abap_true.
      RETURN.
    ENDIF.

    ev_kondm = VALUE #( lt_fields[ fieldname = 'KONDM' ]-value ).
    ev_ktgrm = VALUE #( lt_fields[ fieldname = 'KTGRM' ]-value ).
    ev_mtpos = VALUE #( lt_fields[ fieldname = 'MTPOS' ]-value ).
    ev_tatyp = VALUE #( lt_fields[ fieldname = 'TATYP' ]-value ).
    ev_taxkm = VALUE #( lt_fields[ fieldname = 'TAXKM' ]-value ).

    IF iv_ask_taxim EQ abap_true.
      ev_taxim = VALUE #( lt_fields[ fieldname = 'TAXIM' ]-value ).
    ENDIF.

  ENDMETHOD.


  METHOD ask_valuation_fields.

    DATA(lt_fields) = VALUE ty_sval( ).

    APPEND VALUE #( tabname   = 'MARA'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_source
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MARC'
                    fieldname = 'MATNR'
                    field_obl = abap_false
                    value     = mr_material_selected->matnr_target
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MBEW'
                    fieldname = 'BWKEY'
                    field_obl = abap_false
                    value     = iv_bwkey
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MBEW'
                    fieldname = 'BWTAR'
                    field_obl = abap_false
                    value     = iv_bwtar
                    field_attr = '02' ) TO lt_fields.

    APPEND VALUE #( tabname   = 'MBEW'
                    fieldname = 'STPRS'
                    field_obl = abap_true
                    value     = iv_stprs ) TO lt_fields.

    APPEND VALUE #( tabname   = 'T001'
                    fieldname = 'WAERS'
                    field_obl = abap_false
                    field_attr = '04' ) TO lt_fields.

    popup_get_values(
      EXPORTING
        iv_no_value_check = abap_false
        iv_popup_title    = 'Aşağıdaki bilgileri doldurunuz'
      IMPORTING
        ev_returncode     = DATA(lv_return_code)
      CHANGING
        ct_fields         = lt_fields ).

    ev_aborted = xsdbool( lv_return_code IS NOT INITIAL ).

    IF ev_aborted EQ abap_true.
      RETURN.
    ENDIF.

    ev_stprs = VALUE #( lt_fields[ fieldname = 'STPRS' ]-value ).

  ENDMETHOD.

  METHOD get_reference_material_data.

    DATA(lv_plant) = iv_plant.
    IF lv_plant IS INITIAL.

      SELECT SINGLE werks
        FROM marc
       WHERE matnr EQ @iv_matnr
        INTO @lv_plant.

    ENDIF.

    DATA(lv_whse_no) = iv_whse_no.
    IF lv_whse_no IS INITIAL.
      lv_whse_no = VALUE #( mt_lgnum[ werks = lv_plant ]-lgnum OPTIONAL ).
    ENDIF.

    DATA(lv_salesorg) = iv_salesorg.
    DATA(lv_distr_chan) = iv_distr_chan.
    IF lv_salesorg IS INITIAL OR lv_distr_chan IS INITIAL.

      SELECT SINGLE vkorg, vtweg
        FROM mvke
       WHERE matnr EQ @iv_matnr
        INTO ( @lv_salesorg, @lv_distr_chan ).

    ENDIF.

    DATA(lv_val_area) = iv_val_area.
    DATA(lv_val_type) = iv_val_type.
    IF lv_val_area IS INITIAL OR lv_val_type IS INITIAL.

      SELECT SINGLE bwkey, bwtar
        FROM mbew
       WHERE matnr EQ @iv_matnr
        INTO ( @lv_val_area, @lv_val_type ).

    ENDIF.

    SELECT SINGLE taxim
      FROM mlan
     WHERE matnr EQ @iv_matnr
      INTO @ev_taxim.

    CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        material_long = iv_matnr
        plant         = lv_plant
      IMPORTING
        plantdata     = es_plantdata.

    CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        material_long       = iv_matnr
        whsenumber          = lv_whse_no
      IMPORTING
        warehousenumberdata = es_warehousenumberdata.

    IF es_warehousenumberdata IS INITIAL.
      es_warehousenumberdata-whse_no = lv_whse_no.
    ENDIF.

    DATA(lt_taxclassifications) = VALUE txs_mat_tax_classif_tty( ).
    CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        material_long      = iv_matnr
        salesorg           = lv_salesorg
        distr_chan         = lv_distr_chan
      IMPORTING
        salesdata          = es_salesdata
      TABLES
        taxclassifications = lt_taxclassifications.

    CALL FUNCTION 'BAPI_MATERIAL_GET_ALL' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        material_long = iv_matnr
        val_area      = lv_val_area
        val_type      = lv_val_type
      IMPORTING
        valuationdata = es_valuationdata.

    es_taxclassification = VALUE #( lt_taxclassifications[ 1 ] OPTIONAL ).
  ENDMETHOD.


  METHOD plant_extension.

    DATA(ls_bapi_marcx) = VALUE bapi_marcx( ).
    DATA(ls_bapi_mbewx) = VALUE bapi_mbewx( ).
    DATA(ls_bapi_mvkex) = VALUE bapi_mvkex( ).
    DATA(ls_bapi_mardx) = VALUE bapi_mardx( ).

    fill_bapix( REF #( ls_bapi_marcx ) ).
    fill_bapix( REF #( ls_bapi_mbewx ) ).
    fill_bapix( REF #( ls_bapi_mvkex ) ).
    fill_bapix( REF #( ls_bapi_mardx ) ).

    ls_bapi_marcx-del_flag = ''.
    ls_bapi_mbewx-del_flag = ''.
    ls_bapi_mvkex-del_flag = ''.
    ls_bapi_mardx-del_flag = ''.

    ls_bapi_mbewx-moving_pr   = ''.
    ls_bapi_mbewx-std_pr_pp   = ''.
    ls_bapi_mbewx-std_price   = ''.
    ls_bapi_mbewx-mov_pr_py   = ''.
    ls_bapi_mbewx-std_pr_py   = ''.
    ls_bapi_mbewx-pr_ctrl_pp  = ''.
    ls_bapi_mbewx-mov_pr_pp   = ''.

    CLEAR : ls_bapi_mbewx-future_pr,
            ls_bapi_mbewx-valid_from.

    SELECT matnr, vkorg, vtweg
      FROM mvke
     WHERE matnr = @iv_matnr
      INTO TABLE @DATA(gt_vkorg).

    SELECT matnr, werks
      FROM marc
     WHERE matnr = @iv_matnr
       AND werks = @iv_werks_source
      INTO TABLE @DATA(gt_plant).

    SELECT matnr, bwkey, bwtar
      FROM mbew
     WHERE matnr = @iv_matnr
       AND bwkey = @iv_werks_source
      INTO TABLE @DATA(gt_area).

    SELECT matnr, werks, lgort
      FROM mard
     WHERE matnr = @iv_matnr
       AND werks = @iv_werks_source
      INTO TABLE @DATA(gt_lgort).

    IF gt_lgort IS INITIAL.
      LOOP AT gt_plant ASSIGNING FIELD-SYMBOL(<ls_plant>).
        APPEND VALUE #( matnr = <ls_plant>-matnr
                        werks = <ls_plant>-werks
                        lgort = '' ) TO gt_lgort.
      ENDLOOP.
    ENDIF.

    DATA: ls_clientdata          TYPE bapi_mara_ga,
          ls_plantdata           TYPE bapi_marc_ga,
          ls_storagelocationdata TYPE bapi_mard_ga,
          bapi_return            TYPE bapiret2.

    LOOP AT gt_lgort ASSIGNING FIELD-SYMBOL(<ls_lgort>).

      CLEAR: ls_clientdata, ls_plantdata, ls_storagelocationdata.

      CALL FUNCTION 'BAPI_MATERIAL_GETALL'
        EXPORTING
          material_long       = <ls_lgort>-matnr
          plant               = <ls_lgort>-werks
          storagelocation     = <ls_lgort>-lgort
        IMPORTING
          clientdata          = ls_clientdata
          plantdata           = ls_plantdata
          storagelocationdata = ls_storagelocationdata.

      DATA(bapi_head) = CORRESPONDING bapimathead( ls_clientdata ).
      bapi_head-material        = ls_clientdata-material.
      bapi_head-ind_sector      = ls_clientdata-ind_sector.
      bapi_head-matl_type       = ls_clientdata-matl_type.
      bapi_head-purchase_view   = 'X'.
      bapi_head-mrp_view        = 'X'.
      bapi_head-prt_view        = 'X'.
      bapi_head-storage_view    = 'X'.
      bapi_head-work_sched_view = 'X'.
      bapi_head-quality_view    = 'X'.

      DATA(bapi_marc1) = CORRESPONDING bapi_marc( ls_plantdata ).
      bapi_marc1-plant = iv_werks_target.

      DATA(bapi_marcx) = ls_bapi_marcx.
      bapi_marcx-plant = iv_werks_target.

      DATA(bapi_mard)  = VALUE bapi_mard( ).
      DATA(bapi_mardx) = VALUE bapi_mardx( ).

      IF <ls_lgort>-lgort IS NOT INITIAL.
        bapi_mardx = ls_bapi_mardx.
        bapi_mard  = CORRESPONDING bapi_mard( ls_storagelocationdata ).

        bapi_mard-plant     = iv_werks_target.
        bapi_mard-stge_loc  = ls_storagelocationdata-stge_loc.

        bapi_mardx-plant    = iv_werks_target.
        bapi_mardx-stge_loc = ls_storagelocationdata-stge_loc.
      ELSE.
        bapi_head-storage_view = ''.
      ENDIF.

      CALL FUNCTION 'BAPI_MATERIAL_SAVEDATA' "#EC CI_USAGE_OK[2438131]
        EXPORTING
          headdata             = bapi_head
          plantdata            = bapi_marc1
          plantdatax           = bapi_marcx
          storagelocationdata  = bapi_mard
          storagelocationdatax = bapi_mardx
        IMPORTING
          return               = bapi_return.

      APPEND bapi_return TO et_messages.

      IF bapi_return-type = 'S'.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = 'X'.
      ENDIF.
    ENDLOOP.

    SORT gt_area BY bwtar ASCENDING.

    DATA ls_valuationdata TYPE bapi_mbew_ga.

    LOOP AT gt_area ASSIGNING FIELD-SYMBOL(<ls_area>).

      CLEAR: ls_clientdata, ls_plantdata, ls_storagelocationdata, ls_valuationdata.

      CALL FUNCTION 'BAPI_MATERIAL_GETALL'
        EXPORTING
          material_long       = <ls_area>-matnr
          valuationarea       = <ls_area>-bwkey
          valuationtype       = <ls_area>-bwtar
          plant               = iv_werks_source
        IMPORTING
          clientdata          = ls_clientdata
          plantdata           = ls_plantdata
          storagelocationdata = ls_storagelocationdata
          valuationdata       = ls_valuationdata.

      bapi_head = CORRESPONDING bapimathead( ls_clientdata ).
      bapi_head-material     = ls_clientdata-material.
      bapi_head-ind_sector   = ls_clientdata-ind_sector.
      bapi_head-matl_type    = ls_clientdata-matl_type.
      bapi_head-cost_view    = 'X'.
      bapi_head-account_view = 'X'.

      DATA(bapi_mbew1) = CORRESPONDING bapi_mbew( ls_valuationdata ).
      bapi_mbew1-val_area = iv_werks_target.

      DATA(bapi_mbewx) = ls_bapi_mbewx.
      bapi_mbewx-val_area = iv_werks_target.
      bapi_mbewx-val_type = bapi_mbew1-val_type.

      bapi_marc1 = CORRESPONDING bapi_marc( ls_plantdata ).
      bapi_marc1-plant = iv_werks_target.

      bapi_marcx = ls_bapi_marcx.
      bapi_marcx-plant = iv_werks_target.

      CALL FUNCTION 'BAPI_MATERIAL_SAVEDATA' "#EC CI_USAGE_OK[2438131]
        EXPORTING
          headdata       = bapi_head
          plantdata      = bapi_marc1
          plantdatax     = bapi_marcx
          valuationdata  = bapi_mbew1
          valuationdatax = bapi_mbewx
        IMPORTING
          return         = bapi_return.

      APPEND bapi_return TO et_messages.

      IF bapi_return-type = 'S'.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = 'X'.
      ENDIF.
    ENDLOOP.

    DATA ls_salesdata TYPE bapi_mvke_ga.

    LOOP AT gt_vkorg ASSIGNING FIELD-SYMBOL(<ls_vkorg>).

      CLEAR: ls_clientdata, ls_plantdata, ls_salesdata.

      CALL FUNCTION 'BAPI_MATERIAL_GETALL'
        EXPORTING
          material_long       = <ls_vkorg>-matnr
          plant               = iv_werks_source
          salesorganisation   = <ls_vkorg>-vkorg
          distributionchannel = <ls_vkorg>-vtweg
        IMPORTING
          clientdata          = ls_clientdata
          plantdata           = ls_plantdata
          salesdata           = ls_salesdata.

      bapi_head = CORRESPONDING bapimathead( ls_clientdata ).
      bapi_head-material    = ls_clientdata-material.
      bapi_head-ind_sector  = ls_clientdata-ind_sector.
      bapi_head-matl_type   = ls_clientdata-matl_type.
      bapi_head-sales_view  = 'X'.
      bapi_head-prt_view    = 'X'.

      DATA(bapi_mvke1) = CORRESPONDING bapi_mvke( ls_salesdata ).

      DATA(bapi_mvkex) = ls_bapi_mvkex.
      bapi_mvkex-sales_org  = <ls_vkorg>-vkorg.
      bapi_mvkex-distr_chan = <ls_vkorg>-vtweg.

      bapi_marc1 = CORRESPONDING bapi_marc( ls_plantdata ).
      bapi_marc1-plant = iv_werks_target.

      bapi_marcx = ls_bapi_marcx.
      bapi_marcx-plant = iv_werks_target.

      CALL FUNCTION 'BAPI_MATERIAL_SAVEDATA' "#EC CI_USAGE_OK[2438131]
        EXPORTING
          headdata   = bapi_head
          plantdata  = bapi_marc1
          plantdatax = bapi_marcx
          salesdata  = bapi_mvke1
          salesdatax = bapi_mvkex
        IMPORTING
          return     = bapi_return.

      APPEND bapi_return TO et_messages.

      IF bapi_return-type = 'S'.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = 'X'.
      ENDIF.

    ENDLOOP.

    SELECT *
      FROM qmat
     WHERE matnr = @iv_matnr
       AND werks = @iv_werks_source
      INTO TABLE @DATA(lt_qmat).

    IF sy-subrc IS INITIAL.

      DATA: ls_bapi_qmat   TYPE bapi1001004_qmat,
            bapi_qmat      TYPE STANDARD TABLE OF bapi1001004_qmat,
            lt_bapi_return TYPE STANDARD TABLE OF bapiret2.

      LOOP AT lt_qmat ASSIGNING FIELD-SYMBOL(<ls_qmat>).

        CLEAR ls_bapi_qmat.

        CALL FUNCTION 'MAP2E_QMAT_TO_BAPI1001004_QMAT' "#EC CI_USAGE_OK[2438006]
          EXPORTING
            qmat             = <ls_qmat>
          CHANGING
            bapi1001004_qmat = ls_bapi_qmat.

        ls_bapi_qmat-plant = iv_werks_target.
        APPEND ls_bapi_qmat TO bapi_qmat.

      ENDLOOP.

      CALL FUNCTION 'BAPI_MATINSPCTRL_SAVEREPLICA' "#EC CI_USAGE_OK[2438131]
        TABLES
          return         = lt_bapi_return
          inspectionctrl = bapi_qmat.

      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          wait = 'X'.
    ENDIF.

  ENDMETHOD.

ENDCLASS.
