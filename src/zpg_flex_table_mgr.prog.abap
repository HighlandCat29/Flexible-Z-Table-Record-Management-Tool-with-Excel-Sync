*&---------------------------------------------------------------------*
*& Report ZPG_FLEX_TABLE_MGR
*&---------------------------------------------------------------------*
*& Flexible Z-Table Record Management Tool (SE16N Style)
*& Uses a blank selection screen 2000 as host for OO ALV.
*& This approach is 100% robust on SAP GUI & WebGUI without SE51 screen painter.
*&---------------------------------------------------------------------*
REPORT zpg_flex_table_mgr.

INCLUDE <icon>.

TABLES: sscrfields.

" --- Local Classes Definition (Deferred) ---
CLASS lcl_selection_screen DEFINITION DEFERRED.
CLASS lcl_dynamic_handler DEFINITION DEFERRED.
CLASS lcl_filter_handler DEFINITION DEFERRED.
CLASS lcl_excel_sync DEFINITION DEFERRED.
CLASS lcl_lock_manager DEFINITION DEFERRED.
CLASS lcl_alv_report DEFINITION DEFERRED.
CLASS lcl_event_handler DEFINITION DEFERRED.
CLASS lcl_gsheet_event_handler DEFINITION DEFERRED.
CLASS lcl_audit_logger DEFINITION DEFERRED.
CLASS lcl_field_util DEFINITION DEFERRED.
CLASS lcl_gsheet_payload DEFINITION DEFERRED.
CLASS lcl_alv_style DEFINITION DEFERRED.

" --- Global Data & Object References ---
DATA: go_sel_screen  TYPE REF TO lcl_selection_screen,
      go_dyn_handler TYPE REF TO lcl_dynamic_handler,
      go_filter      TYPE REF TO lcl_filter_handler,
      go_excel_sync  TYPE REF TO lcl_excel_sync,
      go_lock_mgr    TYPE REF TO lcl_lock_manager,
      go_alv_report  TYPE REF TO lcl_alv_report,
      go_audit       TYPE REF TO lcl_audit_logger.

FIELD-SYMBOLS: <dyn_table> TYPE STANDARD TABLE,
               <dyn_wa>    TYPE ANY,
               <dyn_field> TYPE ANY,
               <gsheet_table> TYPE STANDARD TABLE.

DATA: gv_edit_mode   TYPE abap_bool VALUE abap_false,
      gv_has_changes TYPE abap_bool VALUE abap_false,
      gv_upload_preview TYPE abap_bool VALUE abap_false,
      gv_upload_has_errors TYPE abap_bool VALUE abap_false,
      gr_upload_backup TYPE REF TO data,
      gv_dummy_text  TYPE c LENGTH 50.

" Global ALV and Container Objects
DATA: go_grid            TYPE REF TO cl_gui_alv_grid,
      go_sheet_grid      TYPE REF TO cl_gui_alv_grid,
      go_main_splitter   TYPE REF TO cl_gui_splitter_container,
      go_header_splitter TYPE REF TO cl_gui_splitter_container,
      go_top_cont        TYPE REF TO cl_gui_container,
      go_hdr_doc         TYPE REF TO cl_dd_document.

DATA: gr_gsheet_table  TYPE REF TO data,
      gv_gsheet_status TYPE string,
      gv_gsheet_has_instance TYPE abap_bool VALUE abap_false,
      gv_gsheet_can_export   TYPE abap_bool VALUE abap_false,
      gv_action_error  TYPE string.

" --- Selection Screen 1000 (Main Parameters) ---
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_tabnam TYPE tabname OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_maint AS CHECKBOX DEFAULT ' '.
SELECTION-SCREEN END OF BLOCK b2.

" --- Selection Screen 2000 (Container Host Screen) ---
" We add a dummy comment to force WebGUI to render this screen,
" which is then completely covered by our fullscreen ALV.
SELECTION-SCREEN BEGIN OF SCREEN 2000.
  SELECTION-SCREEN COMMENT /1(50) TEST.
SELECTION-SCREEN END OF SCREEN 2000.

" =====================================================================
" CLASS DEFINITIONS
" =====================================================================

CLASS lcl_selection_screen DEFINITION.
  PUBLIC SECTION.
    METHODS: validate_input.
ENDCLASS.

CLASS lcl_dynamic_handler DEFINITION.
  PUBLIC SECTION.
    DATA: mt_fieldcat TYPE lvc_t_fcat.
    METHODS: create_dynamic_table,
             build_fieldcatalog,
             set_edit_mode IMPORTING iv_edit TYPE abap_bool.
ENDCLASS.

CLASS lcl_filter_handler DEFINITION.
  PUBLIC SECTION.
    DATA: mv_sel_id        TYPE dynselid,
          mt_where_clauses TYPE rsds_twhere,
          mv_num_fields    TYPE i,
          mv_initialized   TYPE abap_bool.
    METHODS:
      initialize,
      show_dialog RETURNING VALUE(rv_cancelled) TYPE abap_bool,
      execute_select,
      get_filter_text RETURNING VALUE(rv_text) TYPE string.
ENDCLASS.

CLASS lcl_excel_sync DEFINITION.
  PUBLIC SECTION.
    METHODS: download_excel,
             upload_excel.
ENDCLASS.

CLASS lcl_lock_manager DEFINITION.
  PUBLIC SECTION.
    CONSTANTS: gc_timeout_secs TYPE i VALUE 900,
               gc_table_key_hash TYPE zflex_edit_lock-key_hash VALUE 'TABLE_LOCK', "#EC NOTEXT
               gc_table_field TYPE fieldname VALUE 'TABLE'. "#EC NOTEXT
    METHODS:
      constructor,
      cleanup_expired,
      get_table_lock
        RETURNING VALUE(rs_lock) TYPE zflex_edit_lock,
      is_table_locked_by_other
        RETURNING VALUE(rv_locked) TYPE abap_bool,
      acquire_table_lock
        RETURNING VALUE(rv_success) TYPE abap_bool,
      build_key_text
        IMPORTING is_row TYPE any
        RETURNING VALUE(rv_key_text) TYPE string,
      build_key_hash
        IMPORTING iv_key_text TYPE string
        RETURNING VALUE(rv_key_hash) TYPE zflex_edit_lock-key_hash,
      release_session.
  PRIVATE SECTION.
    DATA: mv_session_id TYPE zflex_edit_lock-session_id.
ENDCLASS.

CLASS lcl_alv_report DEFINITION.
  PUBLIC SECTION.
    METHODS: display_alv,
             save_data,
             accept_upload_preview,
             cancel_upload_preview,
             create_entry,
             delete_entries,
             toggle_edit,
             excel_sync_popup,
             sync_from_cloud,
             push_to_cloud,
             export_to_cloud,
             refresh_header,
             refresh_mode_title,
             build_html_header IMPORTING io_dd TYPE REF TO cl_dd_document.
ENDCLASS.

CLASS lcl_audit_logger DEFINITION.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_audit_hdr,
             client      TYPE mandt,
             change_id   TYPE c LENGTH 32,
             tabname     TYPE tabname,
             action      TYPE c LENGTH 10,
             uname       TYPE syuname,
             change_date TYPE sy-datum,
             change_time TYPE sy-uzeit,
             tstamp      TYPE timestamp,
             row_count   TYPE i,
             source      TYPE c LENGTH 12,
             progname    TYPE syrepid,
             status      TYPE c LENGTH 1,
             description TYPE c LENGTH 200,
           END OF ty_audit_hdr.
    TYPES: BEGIN OF ty_audit_itm,
             client    TYPE mandt,
             change_id TYPE c LENGTH 32,
             item_no   TYPE n LENGTH 6,
             key_hash  TYPE c LENGTH 32,
             key_text  TYPE c LENGTH 500,
             fieldname TYPE fieldname,
             old_value TYPE c LENGTH 255,
             new_value TYPE c LENGTH 255,
           END OF ty_audit_itm.
    METHODS:
      log_access
        IMPORTING
          iv_tabname TYPE tabname
          iv_action  TYPE clike
          iv_source  TYPE clike
          iv_rows    TYPE i
          iv_desc    TYPE clike OPTIONAL,
      log_changes
        IMPORTING
          iv_tabname  TYPE tabname
          iv_action   TYPE clike
          iv_source   TYPE clike
          it_new_data TYPE ANY TABLE
          it_old_data TYPE ANY TABLE OPTIONAL
          iv_desc     TYPE clike OPTIONAL.
  PRIVATE SECTION.
    METHODS:
      generate_change_id
        RETURNING VALUE(rv_id) TYPE ty_audit_hdr-change_id,
      get_timestamp
        RETURNING VALUE(rv_tstamp) TYPE timestamp.
ENDCLASS.

CLASS lcl_event_handler DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: on_toolbar FOR EVENT toolbar OF cl_gui_alv_grid
                     IMPORTING e_object e_interactive,
                   on_user_command FOR EVENT user_command OF cl_gui_alv_grid
                     IMPORTING e_ucomm,
                   on_data_changed FOR EVENT data_changed OF cl_gui_alv_grid
                     IMPORTING er_data_changed.
ENDCLASS.

CLASS lcl_gsheet_event_handler DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: on_toolbar FOR EVENT toolbar OF cl_gui_alv_grid
                     IMPORTING e_object e_interactive,
                   on_user_command FOR EVENT user_command OF cl_gui_alv_grid
                     IMPORTING e_ucomm.
ENDCLASS.

CLASS lcl_field_util DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
      is_technical_field
        IMPORTING iv_fieldname TYPE fieldname
        RETURNING VALUE(rv_is_tech) TYPE abap_bool,
      is_cloud_compare_skip_field
        IMPORTING iv_fieldname TYPE fieldname
        RETURNING VALUE(rv_skip) TYPE abap_bool,
      get_compare_value
        IMPORTING
          is_row  TYPE any
          is_fcat TYPE lvc_s_fcat
        RETURNING VALUE(rv_value) TYPE string,
      build_cloud_compare_key
        IMPORTING
          is_row      TYPE any
          it_key_fcat TYPE lvc_t_fcat
          iv_index    TYPE sy-tabix
        RETURNING VALUE(rv_key) TYPE string,
      fill_generated_technical_keys
        CHANGING cs_row TYPE any.
ENDCLASS.

CLASS lcl_gsheet_payload DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
      json_escape
        IMPORTING iv_input TYPE string
        RETURNING VALUE(rv_output) TYPE string,
      append_json_value
        IMPORTING iv_value TYPE string
        CHANGING cv_json TYPE string,
      column_index_to_a1
        IMPORTING iv_index TYPE i
        RETURNING VALUE(rv_column) TYPE string,
      build_sheet_payload
        IMPORTING it_push TYPE ANY TABLE
        CHANGING
          cv_payload   TYPE string
          cv_range     TYPE string
          cv_row_count TYPE i.
ENDCLASS.

CLASS lcl_alv_style DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
      set_row_status
        IMPORTING iv_status TYPE string
        CHANGING cs_row TYPE any,
      set_row_message
        IMPORTING iv_message TYPE string
        CHANGING cs_row TYPE any,
      mark_cloud_preview_row
        IMPORTING iv_message TYPE string
        CHANGING cs_row TYPE any,
      color_gsheet_row
        IMPORTING
          iv_message TYPE string
          iv_color   TYPE i
        CHANGING cs_row TYPE any,
      color_gsheet_cell
        IMPORTING
          iv_fieldname TYPE fieldname
          iv_color     TYPE i
        CHANGING cs_row TYPE any,
      recolor_row_cells
        IMPORTING
          iv_from_color TYPE i
          iv_to_color   TYPE i
        CHANGING cs_row TYPE any,
      set_gsheet_message
        IMPORTING iv_message TYPE string
        CHANGING cs_row TYPE any.
ENDCLASS.

" =====================================================================
" CLASS IMPLEMENTATIONS
" =====================================================================

CLASS lcl_field_util IMPLEMENTATION.
  METHOD is_technical_field.
    rv_is_tech = abap_false.
    IF iv_fieldname = 'IS_NEW_ROW' "#EC NOTEXT
       OR iv_fieldname = 'CELL_STYLES' "#EC NOTEXT
       OR iv_fieldname = 'ROW_STATUS' "#EC NOTEXT
       OR iv_fieldname = 'CELL_COLORS' "#EC NOTEXT
       OR iv_fieldname = 'LOCK_OWNER' "#EC NOTEXT
       OR iv_fieldname = 'LOCK_FIELD' "#EC NOTEXT
       OR iv_fieldname = 'LOCK_INFO' "#EC NOTEXT
       OR iv_fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
      rv_is_tech = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD is_cloud_compare_skip_field.
    rv_skip = is_technical_field( iv_fieldname ).
    IF rv_skip = abap_true.
      RETURN.
    ENDIF.

    IF iv_fieldname = 'MANDT' "#EC NOTEXT
       OR iv_fieldname = 'CLIENT' "#EC NOTEXT
       OR iv_fieldname = 'UUID' "#EC NOTEXT
       OR iv_fieldname = 'GUID' "#EC NOTEXT
       OR iv_fieldname CS 'UUID' "#EC NOTEXT
       OR iv_fieldname CS 'GUID'. "#EC NOTEXT
      rv_skip = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD get_compare_value.
    CLEAR rv_value.

    ASSIGN COMPONENT is_fcat-fieldname OF STRUCTURE is_row TO FIELD-SYMBOL(<lv_value>).
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    rv_value = |{ <lv_value> }|. "#EC NOTEXT
    CONDENSE rv_value NO-GAPS.

    CASE is_fcat-datatype.
      WHEN 'DATS'. "#EC NOTEXT
        REPLACE ALL OCCURRENCES OF '-' IN rv_value WITH ''. "#EC NOTEXT
      WHEN 'TIMS'. "#EC NOTEXT
        REPLACE ALL OCCURRENCES OF ':' IN rv_value WITH ''. "#EC NOTEXT
    ENDCASE.

    IF rv_value IS NOT INITIAL AND rv_value CO '0123456789'.
      SHIFT rv_value LEFT DELETING LEADING '0'.
      IF rv_value IS INITIAL.
        rv_value = '0'.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD build_cloud_compare_key.
    DATA lv_value TYPE string.

    CLEAR rv_key.

    IF it_key_fcat IS INITIAL.
      rv_key = |ROW={ iv_index }|. "#EC NOTEXT
      RETURN.
    ENDIF.

    LOOP AT it_key_fcat INTO DATA(ls_key_fcat).
      lv_value = get_compare_value(
        is_row  = is_row
        is_fcat = ls_key_fcat ).
      DATA(lv_key_part) = |{ ls_key_fcat-fieldname }={ lv_value }|. "#EC NOTEXT
      CONCATENATE rv_key lv_key_part '|' INTO rv_key.
    ENDLOOP.
  ENDMETHOD.

  METHOD fill_generated_technical_keys.
    FIELD-SYMBOLS <lv_value> TYPE any.

    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_uuid_fcat).
      IF ls_uuid_fcat-fieldname <> 'UUID' "#EC NOTEXT
         AND ls_uuid_fcat-fieldname <> 'GUID' "#EC NOTEXT
         AND ls_uuid_fcat-fieldname NS 'UUID' "#EC NOTEXT
         AND ls_uuid_fcat-fieldname NS 'GUID'. "#EC NOTEXT
        CONTINUE.
      ENDIF.

      ASSIGN COMPONENT ls_uuid_fcat-fieldname OF STRUCTURE cs_row TO <lv_value>.
      IF sy-subrc <> 0 OR <lv_value> IS NOT INITIAL.
        CONTINUE.
      ENDIF.

      TRY.
          IF ls_uuid_fcat-inttype = 'X'. "#EC NOTEXT
            DATA(lv_uuid_x16) = cl_system_uuid=>create_uuid_x16_static( ).
            <lv_value> = lv_uuid_x16.
          ELSE.
            DATA(lv_uuid_c32) = cl_system_uuid=>create_uuid_c32_static( ).
            <lv_value> = lv_uuid_c32.
          ENDIF.
        CATCH cx_uuid_error cx_sy_conversion_error cx_sy_move_cast_error.
          TRY.
              <lv_value> = |{ sy-datum }{ sy-uzeit }{ sy-tabix }|. "#EC NOTEXT
            CATCH cx_root.
          ENDTRY.
      ENDTRY.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_gsheet_payload IMPLEMENTATION.
  METHOD json_escape.
    rv_output = iv_input.
    REPLACE ALL OCCURRENCES OF '\' IN rv_output WITH '\\'. "#EC NOTEXT
    REPLACE ALL OCCURRENCES OF '"' IN rv_output WITH '\"'. "#EC NOTEXT
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN rv_output WITH '\n'. "#EC NOTEXT
  ENDMETHOD.

  METHOD append_json_value.
    DATA(lv_escaped) = json_escape( iv_value ).
    cv_json = |{ cv_json }"{ lv_escaped }"|. "#EC NOTEXT
  ENDMETHOD.

  METHOD column_index_to_a1.
    CONSTANTS lc_alpha TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'. "#EC NOTEXT

    DATA: lv_index TYPE i,
          lv_rem   TYPE i,
          lv_char  TYPE c LENGTH 1.

    CLEAR rv_column.
    lv_index = iv_index.

    WHILE lv_index > 0.
      lv_index = lv_index - 1.
      lv_rem = lv_index MOD 26.
      lv_char = lc_alpha+lv_rem(1).
      rv_column = |{ lv_char }{ rv_column }|. "#EC NOTEXT
      lv_index = lv_index DIV 26.
    ENDWHILE.
  ENDMETHOD.

  METHOD build_sheet_payload.
    DATA: lt_push_fcat TYPE lvc_t_fcat,
          lv_row_json  TYPE string,
          lv_cell      TYPE string,
          lv_last_col  TYPE string,
          lv_last_row  TYPE i.

    FIELD-SYMBOLS: <ls_push_row> TYPE any,
                   <lv_value>    TYPE any.

    CLEAR: cv_payload, cv_range, cv_row_count.

    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_fcat).
      IF lcl_field_util=>is_cloud_compare_skip_field( ls_fcat-fieldname ) = abap_true.
        CONTINUE.
      ENDIF.
      APPEND ls_fcat TO lt_push_fcat.
    ENDLOOP.

    IF lt_push_fcat IS INITIAL.
      gv_action_error = |No fields available to push for { p_tabnam }|. "#EC NOTEXT
      RETURN.
    ENDIF.

    cv_payload = '{"majorDimension":"ROWS","values":['. "#EC NOTEXT

    lv_row_json = '['.
    LOOP AT lt_push_fcat INTO ls_fcat.
      IF sy-tabix > 1.
        CONCATENATE lv_row_json ',' INTO lv_row_json.
      ENDIF.
      lv_cell = ls_fcat-fieldname.
      append_json_value(
        EXPORTING iv_value = lv_cell
        CHANGING  cv_json  = lv_row_json ).
    ENDLOOP.
    CONCATENATE lv_row_json ']' INTO lv_row_json.
    CONCATENATE cv_payload lv_row_json INTO cv_payload.

    LOOP AT it_push ASSIGNING <ls_push_row>.
      CONCATENATE cv_payload ',' INTO cv_payload.
      lv_row_json = '['.

      LOOP AT lt_push_fcat INTO ls_fcat.
        IF sy-tabix > 1.
          CONCATENATE lv_row_json ',' INTO lv_row_json.
        ENDIF.

        CLEAR lv_cell.
        ASSIGN COMPONENT ls_fcat-fieldname OF STRUCTURE <ls_push_row> TO <lv_value>.
        IF sy-subrc = 0.
          lv_cell = |{ <lv_value> }|. "#EC NOTEXT
        ENDIF.
        append_json_value(
          EXPORTING iv_value = lv_cell
          CHANGING  cv_json  = lv_row_json ).
      ENDLOOP.

      CONCATENATE lv_row_json ']' INTO lv_row_json.
      CONCATENATE cv_payload lv_row_json INTO cv_payload.
    ENDLOOP.

    CONCATENATE cv_payload ']}' INTO cv_payload.

    cv_row_count = lines( it_push ).
    lv_last_row = cv_row_count + 1.
    DATA(lv_col_count) = lines( lt_push_fcat ).
    lv_last_col = column_index_to_a1( lv_col_count ).
    cv_range = |A1:{ lv_last_col }{ lv_last_row }|. "#EC NOTEXT
  ENDMETHOD.
ENDCLASS.

CLASS lcl_alv_style IMPLEMENTATION.
  METHOD set_row_status.
    ASSIGN COMPONENT 'ROW_STATUS' OF STRUCTURE cs_row TO FIELD-SYMBOL(<lv_status>). "#EC NOTEXT
    IF sy-subrc = 0.
      CASE iv_status.
        WHEN 'Failed'. "#EC NOTEXT
          <lv_status> = '1'.
        WHEN 'Edit'. "#EC NOTEXT
          <lv_status> = '2'.
        WHEN 'Updated'. "#EC NOTEXT
          <lv_status> = '3'.
        WHEN OTHERS.
          CLEAR <lv_status>.
      ENDCASE.
    ENDIF.
  ENDMETHOD.

  METHOD set_row_message.
    ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE cs_row TO FIELD-SYMBOL(<lv_message>). "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_message> = iv_message.
    ENDIF.
  ENDMETHOD.

  METHOD mark_cloud_preview_row.
    FIELD-SYMBOLS <lv_tech> TYPE any.

    ASSIGN COMPONENT 'MANDT' OF STRUCTURE cs_row TO <lv_tech>. "#EC NOTEXT
    IF sy-subrc = 0 AND <lv_tech> IS INITIAL.
      <lv_tech> = sy-mandt.
    ENDIF.

    ASSIGN COMPONENT 'CLIENT' OF STRUCTURE cs_row TO <lv_tech>. "#EC NOTEXT
    IF sy-subrc = 0 AND <lv_tech> IS INITIAL.
      <lv_tech> = sy-mandt.
    ENDIF.

    ASSIGN COMPONENT 'LOCK_OWNER' OF STRUCTURE cs_row TO <lv_tech>. "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_tech> = sy-uname.
    ENDIF.

    ASSIGN COMPONENT 'LOCK_FIELD' OF STRUCTURE cs_row TO <lv_tech>. "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_tech> = 'CLOUD'. "#EC NOTEXT
    ENDIF.

    ASSIGN COMPONENT 'LOCK_INFO' OF STRUCTURE cs_row TO <lv_tech>. "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_tech> = 'Cloud sync preview by me'. "#EC NOTEXT
    ENDIF.

    set_row_status(
      EXPORTING iv_status = 'Edit'
      CHANGING  cs_row    = cs_row ).
    set_row_message(
      EXPORTING iv_message = iv_message
      CHANGING  cs_row     = cs_row ).
  ENDMETHOD.

  METHOD color_gsheet_row.
    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_row_fcat).
      IF lcl_field_util=>is_technical_field( ls_row_fcat-fieldname ) = abap_true
         OR ls_row_fcat-fieldname = 'MANDT' "#EC NOTEXT
         OR ls_row_fcat-fieldname = 'CLIENT'. "#EC NOTEXT
        CONTINUE.
      ENDIF.
      color_gsheet_cell(
        EXPORTING
          iv_fieldname = ls_row_fcat-fieldname
          iv_color     = iv_color
        CHANGING
          cs_row       = cs_row ).
    ENDLOOP.

    set_gsheet_message(
      EXPORTING iv_message = iv_message
      CHANGING  cs_row     = cs_row ).
  ENDMETHOD.

  METHOD color_gsheet_cell.
    FIELD-SYMBOLS <lt_colors> TYPE lvc_t_scol.

    ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE cs_row TO <lt_colors>. "#EC NOTEXT
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    DELETE <lt_colors> WHERE fname = iv_fieldname.

    DATA ls_color TYPE lvc_s_scol.
    ls_color-fname = iv_fieldname.
    ls_color-color-col = iv_color.
    ls_color-color-int = 1.
    ls_color-color-inv = 0.
    APPEND ls_color TO <lt_colors>.
  ENDMETHOD.

  METHOD recolor_row_cells.
    FIELD-SYMBOLS <lt_colors> TYPE lvc_t_scol.

    ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE cs_row TO <lt_colors>. "#EC NOTEXT
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    LOOP AT <lt_colors> ASSIGNING FIELD-SYMBOL(<ls_color>)
      WHERE color-col = iv_from_color.
      <ls_color>-color-col = iv_to_color.
    ENDLOOP.
  ENDMETHOD.

  METHOD set_gsheet_message.
    ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE cs_row TO FIELD-SYMBOL(<lv_message>). "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_message> = iv_message.
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_selection_screen IMPLEMENTATION.
  METHOD validate_input.
    IF p_tabnam(1) <> 'Z' AND p_tabnam(1) <> 'Y'. "#EC NOTEXT
      MESSAGE e001(zflex_msg).
    ENDIF.
    DATA: lv_tabname TYPE dd02l-tabname.
    SELECT SINGLE tabname FROM dd02l INTO @lv_tabname
      WHERE tabname = @p_tabnam AND tabclass = 'TRANSP' AND as4local = 'A'. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE e002(zflex_msg).
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '03' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE e003(zflex_msg) WITH p_tabnam.
    ENDIF.

    IF p_maint = abap_true.
      AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
               ID 'ACTVT' FIELD '02' "#EC NOTEXT
               ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
      IF sy-subrc <> 0.
        MESSAGE e004(zflex_msg) WITH p_tabnam.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_dynamic_handler IMPLEMENTATION.
  METHOD create_dynamic_table.
    DATA: lo_struct_desc TYPE REF TO cl_abap_structdescr,
          lo_table_desc  TYPE REF TO cl_abap_tabledescr,
          lt_comp        TYPE cl_abap_structdescr=>component_table,
          ls_comp        LIKE LINE OF lt_comp.

    lo_struct_desc ?= cl_abap_typedescr=>describe_by_name( p_tabnam ).
    lt_comp = lo_struct_desc->get_components( ).

    ls_comp-name = 'IS_NEW_ROW'. "#EC NOTEXT
    ls_comp-type ?= cl_abap_typedescr=>describe_by_name( 'ABAP_BOOL' ). "#EC NOTEXT
    APPEND ls_comp TO lt_comp.

    ls_comp-name = 'CELL_STYLES'. "#EC NOTEXT
    ls_comp-type ?= cl_abap_typedescr=>describe_by_name( 'LVC_T_STYL' ). "#EC NOTEXT
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'ROW_STATUS'. "#EC NOTEXT
    ls_comp-type = cl_abap_elemdescr=>get_c( 1 ).
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'CELL_COLORS'. "#EC NOTEXT
    ls_comp-type ?= cl_abap_typedescr=>describe_by_name( 'LVC_T_SCOL' ). "#EC NOTEXT
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'LOCK_OWNER'. "#EC NOTEXT
    ls_comp-type = cl_abap_elemdescr=>get_c( 12 ).
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'LOCK_FIELD'. "#EC NOTEXT
    ls_comp-type = cl_abap_elemdescr=>get_c( 30 ).
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'LOCK_INFO'. "#EC NOTEXT
    ls_comp-type = cl_abap_elemdescr=>get_c( 80 ).
    APPEND ls_comp TO lt_comp.

    CLEAR ls_comp.
    ls_comp-name = 'ROW_MESSAGE'. "#EC NOTEXT
    ls_comp-type = cl_abap_elemdescr=>get_c( 200 ).
    APPEND ls_comp TO lt_comp.

    lo_struct_desc = cl_abap_structdescr=>create( p_components = lt_comp ).
    lo_table_desc = cl_abap_tabledescr=>create( p_line_type = lo_struct_desc ).

    DATA: dref_table TYPE REF TO data,
          dref_wa    TYPE REF TO data.
    CREATE DATA dref_table TYPE HANDLE lo_table_desc.
    CREATE DATA dref_wa TYPE HANDLE lo_struct_desc.

    ASSIGN dref_table->* TO <dyn_table>.
    ASSIGN dref_wa->* TO <dyn_wa>.
  ENDMETHOD.

  METHOD build_fieldcatalog.
    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE' "#EC NOTEXT
      EXPORTING
        i_structure_name       = p_tabnam
      CHANGING
        ct_fieldcat            = mt_fieldcat
      EXCEPTIONS
        inconsistent_interface = 1
        program_error          = 2
        OTHERS                 = 3.
    IF sy-subrc <> 0.
      MESSAGE e005(zflex_msg) WITH p_tabnam.
    ENDIF.
    set_edit_mode( gv_edit_mode ).

    LOOP AT mt_fieldcat ASSIGNING FIELD-SYMBOL(<ls_pos_fcat>).
      IF <ls_pos_fcat>-col_pos IS INITIAL.
        <ls_pos_fcat>-col_pos = sy-tabix + 10.
      ELSE.
        <ls_pos_fcat>-col_pos = <ls_pos_fcat>-col_pos + 10.
      ENDIF.
      IF <ls_pos_fcat>-fieldname = 'MANDT' "#EC NOTEXT
         OR <ls_pos_fcat>-fieldname = 'CLIENT'. "#EC NOTEXT
        <ls_pos_fcat>-tech = 'X'. "#EC NOTEXT
        <ls_pos_fcat>-no_out = 'X'. "#EC NOTEXT
      ELSEIF <ls_pos_fcat>-key = 'X'. "#EC NOTEXT
        <ls_pos_fcat>-emphasize = 'C310'. "#EC NOTEXT
      ENDIF.
    ENDLOOP.

    DATA: ls_lock_fcat TYPE lvc_s_fcat.

    CLEAR ls_lock_fcat.
    ls_lock_fcat-fieldname = 'ROW_STATUS'. "#EC NOTEXT
    ls_lock_fcat-coltext   = 'Status'. "#EC NOTEXT
    ls_lock_fcat-scrtext_l = 'Status'. "#EC NOTEXT
    ls_lock_fcat-scrtext_m = 'Status'. "#EC NOTEXT
    ls_lock_fcat-scrtext_s = 'Status'. "#EC NOTEXT
    ls_lock_fcat-outputlen = 6.
    ls_lock_fcat-col_pos   = 1.
    ls_lock_fcat-edit      = space.
    IF gv_edit_mode = abap_false.
      ls_lock_fcat-no_out = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_lock_fcat TO mt_fieldcat.

    CLEAR ls_lock_fcat.
    ls_lock_fcat-fieldname = 'LOCK_OWNER'. "#EC NOTEXT
    ls_lock_fcat-coltext   = 'Lock User'. "#EC NOTEXT
    ls_lock_fcat-scrtext_l = 'Lock User'. "#EC NOTEXT
    ls_lock_fcat-scrtext_m = 'Lock User'. "#EC NOTEXT
    ls_lock_fcat-scrtext_s = 'User'. "#EC NOTEXT
    ls_lock_fcat-outputlen = 12.
    ls_lock_fcat-edit      = space.
    IF gv_edit_mode = abap_false.
      ls_lock_fcat-no_out = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_lock_fcat TO mt_fieldcat.

    CLEAR ls_lock_fcat.
    ls_lock_fcat-fieldname = 'LOCK_FIELD'. "#EC NOTEXT
    ls_lock_fcat-coltext   = 'Last Field'. "#EC NOTEXT
    ls_lock_fcat-scrtext_l = 'Last Field'. "#EC NOTEXT
    ls_lock_fcat-scrtext_m = 'Last Field'. "#EC NOTEXT
    ls_lock_fcat-scrtext_s = 'Field'. "#EC NOTEXT
    ls_lock_fcat-outputlen = 20.
    ls_lock_fcat-edit      = space.
    ls_lock_fcat-tech      = 'X'. "#EC NOTEXT
    ls_lock_fcat-no_out    = 'X'. "#EC NOTEXT
    APPEND ls_lock_fcat TO mt_fieldcat.

    CLEAR ls_lock_fcat.
    ls_lock_fcat-fieldname = 'LOCK_INFO'. "#EC NOTEXT
    ls_lock_fcat-coltext   = 'Edit Lock'. "#EC NOTEXT
    ls_lock_fcat-scrtext_l = 'Edit Lock'. "#EC NOTEXT
    ls_lock_fcat-scrtext_m = 'Edit Lock'. "#EC NOTEXT
    ls_lock_fcat-scrtext_s = 'Lock'. "#EC NOTEXT
    ls_lock_fcat-outputlen = 45.
    ls_lock_fcat-edit      = space.
    ls_lock_fcat-tech      = 'X'. "#EC NOTEXT
    ls_lock_fcat-no_out    = 'X'. "#EC NOTEXT
    APPEND ls_lock_fcat TO mt_fieldcat.

    CLEAR ls_lock_fcat.
    ls_lock_fcat-fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
    ls_lock_fcat-coltext   = 'Message'. "#EC NOTEXT
    ls_lock_fcat-scrtext_l = 'Message'. "#EC NOTEXT
    ls_lock_fcat-scrtext_m = 'Message'. "#EC NOTEXT
    ls_lock_fcat-scrtext_s = 'Msg'. "#EC NOTEXT
    ls_lock_fcat-outputlen = 60.
    ls_lock_fcat-edit      = space.
    IF gv_edit_mode = abap_false.
      ls_lock_fcat-no_out = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_lock_fcat TO mt_fieldcat.
  ENDMETHOD.

  METHOD set_edit_mode.
    FIELD-SYMBOLS: <lvc_fcat> TYPE lvc_s_fcat.
    LOOP AT mt_fieldcat ASSIGNING <lvc_fcat>.
*      <lvc_fcat>-valexi     = abap_true.
      <lvc_fcat>-f4availabl = abap_true.
      IF <lvc_fcat>-fieldname = 'MANDT' "#EC NOTEXT
        OR <lvc_fcat>-fieldname = 'CLIENT'. "#EC NOTEXT
        <lvc_fcat>-edit = ' '.
        <lvc_fcat>-tech = 'X'. "#EC NOTEXT
        <lvc_fcat>-no_out = 'X'. "#EC NOTEXT
      ELSEIF <lvc_fcat>-fieldname = 'ROW_STATUS' "#EC NOTEXT
        OR <lvc_fcat>-fieldname = 'LOCK_OWNER' "#EC NOTEXT
        OR <lvc_fcat>-fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
        <lvc_fcat>-edit = ' '.
        IF iv_edit = abap_true.
          <lvc_fcat>-no_out = ' '.
        ELSE.
          <lvc_fcat>-no_out = 'X'. "#EC NOTEXT
        ENDIF.
      ELSEIF <lvc_fcat>-fieldname = 'LOCK_FIELD' "#EC NOTEXT
        OR <lvc_fcat>-fieldname = 'LOCK_INFO'. "#EC NOTEXT
        <lvc_fcat>-edit = ' '.
        <lvc_fcat>-tech = 'X'. "#EC NOTEXT
        <lvc_fcat>-no_out = 'X'. "#EC NOTEXT
      ELSEIF iv_edit = abap_true.
        <lvc_fcat>-edit = 'X'. "#EC NOTEXT
        <lvc_fcat>-no_out = ' '.
      ELSE.
        <lvc_fcat>-edit = ' '.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_filter_handler IMPLEMENTATION.
  METHOD initialize.
    DATA: lt_tables TYPE TABLE OF rsdstabs,
          ls_table  TYPE rsdstabs.
    ls_table-prim_tab = p_tabnam.
    APPEND ls_table TO lt_tables.
    CALL FUNCTION 'FREE_SELECTIONS_INIT' "#EC NOTEXT
      EXPORTING
        kind         = 'T' "#EC NOTEXT
      IMPORTING
        selection_id = mv_sel_id
      TABLES
        tables_tab   = lt_tables
      EXCEPTIONS
        OTHERS       = 1.
    IF sy-subrc = 0.
      mv_initialized = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD show_dialog.
    DATA: lt_fields TYPE TABLE OF rsdsfields,
          lv_title  TYPE sy-title.
    rv_cancelled = abap_true.
    IF mv_initialized = abap_false.
      RETURN.
    ENDIF.
    lv_title = |Selection Criteria - { p_tabnam }|. "#EC NOTEXT
    CALL FUNCTION 'FREE_SELECTIONS_DIALOG' "#EC NOTEXT
      EXPORTING
        selection_id            = mv_sel_id
        title                   = lv_title
        as_window               = 'X' "#EC NOTEXT
      IMPORTING
        where_clauses           = mt_where_clauses
        number_of_active_fields = mv_num_fields
      TABLES
        fields_tab              = lt_fields
      EXCEPTIONS
        internal_error          = 1
        no_action               = 2
        selid_not_found         = 3
        illegal_status          = 4
        OTHERS                  = 5.
    IF sy-subrc = 0.
      rv_cancelled = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD execute_select.
    DATA: lt_where_tab TYPE rsds_where_tab.
    CLEAR <dyn_table>.
    " Extract WHERE clause for our table
    READ TABLE mt_where_clauses INTO DATA(ls_where)
      WITH KEY tablename = p_tabnam.
    IF sy-subrc = 0 AND ls_where-where_tab IS NOT INITIAL.
      SELECT * FROM (p_tabnam) INTO CORRESPONDING FIELDS OF TABLE @<dyn_table>
        WHERE (ls_where-where_tab).
    ELSE.
      SELECT * FROM (p_tabnam) INTO CORRESPONDING FIELDS OF TABLE @<dyn_table>.
    ENDIF.

    PERFORM lock_existing_keys USING <dyn_table>.
    IF go_grid IS BOUND.
      go_grid->refresh_table_display( ).
    ENDIF.

    " === AUDIT LOG: DISPLAY ===
    IF go_audit IS BOUND.
      go_audit->log_access(
        iv_tabname = p_tabnam
        iv_action  = 'DISPLAY' "#EC NOTEXT
        iv_source  = 'ALV' "#EC NOTEXT
        iv_rows    = lines( <dyn_table> )
        iv_desc    = |Viewed { lines( <dyn_table> ) } rows| ). "#EC NOTEXT
    ENDIF.
  ENDMETHOD.

  METHOD get_filter_text.
    IF mv_num_fields > 0.
      rv_text = |{ mv_num_fields } filter(s) active|. "#EC NOTEXT
    ELSE.
      rv_text = 'No filters'. "#EC NOTEXT
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_excel_sync IMPLEMENTATION.
  METHOD download_excel.
    DATA: lv_filename TYPE string,
          lv_path     TYPE string,
          lv_fullpath TYPE string.
    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
        window_title      = |Download data from { p_tabnam }| "#EC NOTEXT
        default_extension = 'XLSX' "#EC NOTEXT
        default_file_name = |{ p_tabnam }| "#EC NOTEXT
        file_filter       = 'Excel Workbook (*.XLSX)|*.XLSX' "#EC NOTEXT
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = lv_fullpath
      EXCEPTIONS
        OTHERS            = 1 ).
    IF sy-subrc <> 0 OR lv_fullpath IS INITIAL.
      RETURN.
    ENDIF.

    DATA: dref_flat_tab TYPE REF TO data,
          lv_xlsx       TYPE xstring,
          lt_bin        TYPE solix_tab,
          lv_size       TYPE i.
    FIELD-SYMBOLS: <flat_tab> TYPE STANDARD TABLE,
                   <flat_wa>  TYPE ANY.

    CREATE DATA dref_flat_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_flat_tab->* TO <flat_tab>.

    LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
      APPEND INITIAL LINE TO <flat_tab> ASSIGNING <flat_wa>.
      MOVE-CORRESPONDING <dyn_wa> TO <flat_wa>.
    ENDLOOP.

    TRY.
        lv_xlsx = cl_fdt_xl_spreadsheet=>if_fdt_doc_spreadsheet~create_document(
          name          = |{ p_tabnam }.xlsx| "#EC NOTEXT
          itab          = dref_flat_tab
          iv_call_type  = if_fdt_doc_spreadsheet=>gc_call_dec_table
          iv_sheet_name = CONV string( p_tabnam ) ).
      CATCH cx_fdt_excel_core INTO DATA(lx_xlsx).
        MESSAGE lx_xlsx->get_text( ) TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
        RETURN.
    ENDTRY.

    lt_bin = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xlsx ).
    lv_size = xstrlen( lv_xlsx ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize = lv_size
        filename     = lv_fullpath
        filetype     = 'BIN' "#EC NOTEXT
      CHANGING
        data_tab = lt_bin
      EXCEPTIONS
        OTHERS   = 1 ).
    IF sy-subrc = 0.
      " === AUDIT LOG: DOWNLOAD ===
      IF go_audit IS BOUND.
        go_audit->log_access(
          iv_tabname = p_tabnam
          iv_action  = 'DOWNLOAD' "#EC NOTEXT
          iv_source  = 'EXCEL_DL' "#EC NOTEXT
          iv_rows    = lines( <dyn_table> )
          iv_desc    = |Downloaded { lines( <dyn_table> ) } rows to XLSX| ). "#EC NOTEXT
      ENDIF.
      DATA(lv_downloaded_rows) = lines( <dyn_table> ).
      MESSAGE s010(zflex_msg) WITH lv_downloaded_rows.
    ELSE.
      MESSAGE s011(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD upload_excel.
    TYPES: BEGIN OF ty_cell_error,
             row       TYPE i,
             fieldname TYPE fieldname,
             message   TYPE string,
           END OF ty_cell_error.

    DATA: lt_file_table TYPE filetable,
          lv_rc         TYPE i,
          lv_preview_row TYPE i,
          lv_row_error TYPE abap_bool,
          lv_key_error TYPE abap_bool,
          lv_key_text TYPE string,
          lv_key_hash TYPE zflex_edit_lock-key_hash,
          lv_error TYPE string,
          lv_header_field TYPE fieldname,
          lt_headers TYPE STANDARD TABLE OF fieldname WITH DEFAULT KEY,
          lt_seen_keys TYPE HASHED TABLE OF zflex_edit_lock-key_hash WITH UNIQUE KEY table_line,
          lt_cell_errors TYPE STANDARD TABLE OF ty_cell_error,
          ls_cell_error TYPE ty_cell_error,
          ls_color TYPE lvc_s_scol.
    FIELD-SYMBOLS: <backup_tab> TYPE STANDARD TABLE,
                   <lv_field> TYPE ANY,
                   <lv_tech> TYPE ANY,
                   <lt_colors> TYPE lvc_t_scol.

    IF gv_edit_mode = abap_false.
      MESSAGE s014(zflex_msg) WITH 'uploading data' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_upload_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s032(zflex_msg) WITH p_tabnam ls_upload_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s004(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        window_title      = |Upload data to { p_tabnam }| "#EC NOTEXT
        default_extension = 'XLSX' "#EC NOTEXT
        file_filter       = 'Excel (*.XLSX;*.XLS)|*.XLSX;*.XLS' "#EC NOTEXT
      CHANGING
        file_table        = lt_file_table
        rc                = lv_rc
      EXCEPTIONS
        OTHERS            = 1 ).
    IF sy-subrc <> 0 OR lv_rc <> 1.
      RETURN.
    ENDIF.
    READ TABLE lt_file_table INTO DATA(ls_file) INDEX 1.
    DATA(lv_filepath) = CONV rlgrap-filename( ls_file-filename ).
    DATA(lv_file_string) = CONV string( ls_file-filename ).
    DATA(lv_file_upper) = lv_file_string.
    TRANSLATE lv_file_upper TO UPPER CASE.
    DATA: lt_raw_data TYPE truxs_t_text_data.

    DATA: dref_flat_tab TYPE REF TO data.
    FIELD-SYMBOLS: <flat_tab> TYPE STANDARD TABLE,
                   <flat_wa>  TYPE ANY.
    CREATE DATA dref_flat_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_flat_tab->* TO <flat_tab>.

    IF lv_file_upper CP '*.XLSX'. "#EC NOTEXT
      DATA: lt_bin             TYPE solix_tab,
            lv_filelength      TYPE i,
            lv_xlsx            TYPE xstring,
            lo_excel           TYPE REF TO cl_fdt_xl_spreadsheet,
            lt_worksheet_names TYPE if_fdt_doc_spreadsheet=>t_worksheet_names,
            lv_sheet_name      TYPE string,
            dref_xlsx_tab      TYPE REF TO data.
      FIELD-SYMBOLS: <xlsx_tab>  TYPE STANDARD TABLE,
                     <xlsx_wa>   TYPE ANY,
                     <xlsx_cell> TYPE ANY,
                     <lv_comp>   TYPE ANY.

      cl_gui_frontend_services=>gui_upload(
        EXPORTING
          filename   = lv_file_string
          filetype   = 'BIN' "#EC NOTEXT
        IMPORTING
          filelength = lv_filelength
        CHANGING
          data_tab   = lt_bin
        EXCEPTIONS
          OTHERS     = 1 ).
      IF sy-subrc <> 0.
        MESSAGE s015(zflex_msg) DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      lv_xlsx = cl_bcs_convert=>solix_to_xstring(
        it_solix = lt_bin
        iv_size  = lv_filelength ).

      TRY.
          CREATE OBJECT lo_excel
            EXPORTING
              document_name = lv_file_string
              xdocument     = lv_xlsx
              mime_type     = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'. "#EC NOTEXT

          lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
            IMPORTING
              worksheet_names = lt_worksheet_names ).
          READ TABLE lt_worksheet_names INTO lv_sheet_name INDEX 1.
          IF sy-subrc <> 0 OR lv_sheet_name IS INITIAL.
            MESSAGE s016(zflex_msg) DISPLAY LIKE 'E'.
            RETURN.
          ENDIF.

          dref_xlsx_tab = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
            worksheet_name = lv_sheet_name ).
        CATCH cx_fdt_excel_core INTO DATA(lx_excel).
          MESSAGE lx_excel->get_text( ) TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
          RETURN.
      ENDTRY.

      ASSIGN dref_xlsx_tab->* TO <xlsx_tab>.
      IF sy-subrc <> 0 OR <xlsx_tab> IS INITIAL.
        MESSAGE s017(zflex_msg) DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      LOOP AT <xlsx_tab> ASSIGNING <xlsx_wa>.
        DATA(lv_xlsx_row) = sy-tabix.
        IF sy-tabix = 1.
          DO.
            ASSIGN COMPONENT sy-index OF STRUCTURE <xlsx_wa> TO <xlsx_cell>.
            IF sy-subrc <> 0.
              EXIT.
            ENDIF.
            DATA(lv_header) = |{ <xlsx_cell> }|. "#EC NOTEXT
            TRANSLATE lv_header TO UPPER CASE.
            CONDENSE lv_header.
            IF lv_header IS INITIAL.
              EXIT.
            ENDIF.
            lv_header_field = lv_header.
            IF lv_header = 'IS_NEW_ROW' "#EC NOTEXT
              OR lv_header = 'CELL_STYLES' "#EC NOTEXT
              OR lv_header = 'ROW_STATUS' "#EC NOTEXT
              OR lv_header = 'CELL_COLORS' "#EC NOTEXT
              OR lv_header = 'LOCK_OWNER' "#EC NOTEXT
              OR lv_header = 'LOCK_FIELD' "#EC NOTEXT
              OR lv_header = 'LOCK_INFO' "#EC NOTEXT
              OR lv_header = 'ROW_MESSAGE'. "#EC NOTEXT
              CONTINUE.
            ENDIF.
            READ TABLE lt_headers TRANSPORTING NO FIELDS
              WITH KEY table_line = lv_header_field.
            IF sy-subrc = 0.
              lv_error = |Duplicate XLSX header field { lv_header }.|. "#EC NOTEXT
              MESSAGE lv_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
              RETURN.
            ENDIF.
            READ TABLE go_dyn_handler->mt_fieldcat TRANSPORTING NO FIELDS
              WITH KEY fieldname = lv_header.
            IF sy-subrc <> 0.
              lv_error = |Unknown XLSX header field { lv_header }.|. "#EC NOTEXT
              MESSAGE lv_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
              RETURN.
            ENDIF.
            APPEND lv_header_field TO lt_headers.
          ENDDO.
          CONTINUE.
        ENDIF.

        DATA(lv_row_has_data) = abap_false.
        APPEND INITIAL LINE TO <flat_tab> ASSIGNING <flat_wa>.
        LOOP AT lt_headers INTO DATA(lv_fieldname).
          DATA(lv_col) = sy-tabix.
          ASSIGN COMPONENT lv_col OF STRUCTURE <xlsx_wa> TO <xlsx_cell>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.
          ASSIGN COMPONENT lv_fieldname OF STRUCTURE <flat_wa> TO <lv_comp>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.
          DATA(lv_cell_value) = |{ <xlsx_cell> }|. "#EC NOTEXT
          IF lv_cell_value IS NOT INITIAL.
            lv_row_has_data = abap_true.
          ENDIF.
          TRY.
              <lv_comp> = lv_cell_value.
            CATCH cx_root INTO DATA(lx_conversion).
              CLEAR ls_cell_error.
              ls_cell_error-row = lines( <flat_tab> ).
              ls_cell_error-fieldname = lv_fieldname.
              ls_cell_error-message = lx_conversion->get_text( ).
              APPEND ls_cell_error TO lt_cell_errors.
              gv_upload_has_errors = abap_true.
          ENDTRY.
        ENDLOOP.

        IF lv_row_has_data = abap_false.
          DATA(lv_last_index) = lines( <flat_tab> ).
          DELETE <flat_tab> INDEX lv_last_index.
        ENDIF.
      ENDLOOP.
    ELSE.
      CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP' "#EC NOTEXT
        EXPORTING
          i_line_header        = 'X' "#EC NOTEXT
          i_tab_raw_data       = lt_raw_data
          i_filename           = lv_filepath
        TABLES
          i_tab_converted_data = <flat_tab>
        EXCEPTIONS
          conversion_failed    = 1
          OTHERS               = 2.
      IF sy-subrc <> 0.
        MESSAGE s024(zflex_msg) DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_xls_fcat)
        WHERE key = 'X' "#EC NOTEXT
          AND fieldname <> 'MANDT' "#EC NOTEXT
          AND fieldname <> 'CLIENT'. "#EC NOTEXT
        APPEND ls_xls_fcat-fieldname TO lt_headers.
      ENDLOOP.
    ENDIF.

    IF <flat_tab> IS INITIAL.
      MESSAGE s017(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF lt_headers IS INITIAL.
      MESSAGE s018(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_key_fcat)
      WHERE key = 'X' "#EC NOTEXT
        AND fieldname <> 'MANDT' "#EC NOTEXT
        AND fieldname <> 'CLIENT'. "#EC NOTEXT
      READ TABLE lt_headers TRANSPORTING NO FIELDS
        WITH KEY table_line = ls_key_fcat-fieldname.
      IF sy-subrc <> 0.
        lv_error = |Upload template missing key field { ls_key_fcat-fieldname }.|. "#EC NOTEXT
        MESSAGE lv_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
        RETURN.
      ENDIF.
    ENDLOOP.

    CREATE DATA gr_upload_backup LIKE <dyn_table>.
    ASSIGN gr_upload_backup->* TO <backup_tab>.
    <backup_tab> = <dyn_table>.

    CLEAR <dyn_table>.
    CLEAR gv_upload_has_errors.

    LOOP AT <flat_tab> ASSIGNING <flat_wa>.
      lv_preview_row = sy-tabix.
      lv_row_error = abap_false.
      APPEND INITIAL LINE TO <dyn_table> ASSIGNING <dyn_wa>.
      MOVE-CORRESPONDING <flat_wa> TO <dyn_wa>.

      ASSIGN COMPONENT 'MANDT' OF STRUCTURE <dyn_wa> TO <lv_field>. "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_field> IS INITIAL.
        <lv_field> = sy-mandt.
      ENDIF.
      ASSIGN COMPONENT 'CLIENT' OF STRUCTURE <dyn_wa> TO <lv_field>. "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_field> IS INITIAL.
        <lv_field> = sy-mandt.
      ENDIF.

      ASSIGN COMPONENT 'LOCK_OWNER' OF STRUCTURE <dyn_wa> TO <lv_tech>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_tech> = sy-uname.
      ENDIF.
      ASSIGN COMPONENT 'LOCK_FIELD' OF STRUCTURE <dyn_wa> TO <lv_tech>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_tech> = 'UPLOAD'. "#EC NOTEXT
      ENDIF.
      ASSIGN COMPONENT 'LOCK_INFO' OF STRUCTURE <dyn_wa> TO <lv_tech>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_tech> = 'Upload preview by me'. "#EC NOTEXT
      ENDIF.

      UNASSIGN <lt_colors>.
      ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE <dyn_wa> TO <lt_colors>. "#EC NOTEXT
      IF sy-subrc = 0.
        CLEAR <lt_colors>.
      ENDIF.

      LOOP AT lt_cell_errors INTO ls_cell_error WHERE row = lv_preview_row.
        lv_row_error = abap_true.
        lv_error = |Invalid { ls_cell_error-fieldname }: { ls_cell_error-message }|. "#EC NOTEXT
        PERFORM set_row_message USING <dyn_wa> lv_error.
        IF <lt_colors> IS ASSIGNED.
          CLEAR ls_color.
          ls_color-fname = ls_cell_error-fieldname.
          ls_color-color-col = 6.
          ls_color-color-int = 1.
          ls_color-color-inv = 0.
          APPEND ls_color TO <lt_colors>.
        ENDIF.
      ENDLOOP.

      lv_key_error = abap_false.
      LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_row_key_fcat)
        WHERE key = 'X'. "#EC NOTEXT
        ASSIGN COMPONENT ls_row_key_fcat-fieldname OF STRUCTURE <dyn_wa> TO <lv_field>.
        IF sy-subrc <> 0 OR <lv_field> IS INITIAL.
          IF ls_row_key_fcat-fieldname <> 'MANDT' "#EC NOTEXT
             AND ls_row_key_fcat-fieldname <> 'CLIENT'. "#EC NOTEXT
            lv_key_error = abap_true.
            lv_row_error = abap_true.
            lv_error = |Missing key field { ls_row_key_fcat-fieldname }|. "#EC NOTEXT
            PERFORM set_row_message USING <dyn_wa> lv_error.
            IF <lt_colors> IS ASSIGNED.
              CLEAR ls_color.
              ls_color-fname = ls_row_key_fcat-fieldname.
              ls_color-color-col = 6.
              ls_color-color-int = 1.
              ls_color-color-inv = 0.
              APPEND ls_color TO <lt_colors>.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDLOOP.

      lv_key_text = go_lock_mgr->build_key_text( <dyn_wa> ).
      lv_key_hash = go_lock_mgr->build_key_hash( lv_key_text ).
      IF lv_key_hash IS INITIAL.
        lv_row_error = abap_true.
        PERFORM set_row_message USING <dyn_wa> 'Cannot build key for upload row'. "#EC NOTEXT
      ELSE.
        READ TABLE lt_seen_keys TRANSPORTING NO FIELDS
          WITH TABLE KEY table_line = lv_key_hash.
        IF sy-subrc = 0.
          lv_row_error = abap_true.
          PERFORM set_row_message USING <dyn_wa> 'Duplicate key in upload file'. "#EC NOTEXT
        ELSE.
          INSERT lv_key_hash INTO TABLE lt_seen_keys.
        ENDIF.
      ENDIF.

      IF lv_row_error = abap_true OR lv_key_error = abap_true.
        gv_upload_has_errors = abap_true.
        PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
      ELSE.
        PERFORM set_row_status USING <dyn_wa> 'Edit'. "#EC NOTEXT
      ENDIF.
    ENDLOOP.

    gv_upload_preview = abap_true.
    gv_has_changes = abap_true.

    IF go_grid IS BOUND.
      go_grid->refresh_table_display( ).
      go_grid->set_toolbar_interactive( ).
      cl_gui_cfw=>flush( ).
    ENDIF.
    IF go_alv_report IS BOUND.
      go_alv_report->refresh_header( ).
    ENDIF.

    IF gv_upload_has_errors = abap_true.
      MESSAGE s019(zflex_msg) DISPLAY LIKE 'E'.
    ELSE.
      DATA(lv_preview_rows) = lines( <dyn_table> ).
      MESSAGE s020(zflex_msg) WITH lv_preview_rows.
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_lock_manager IMPLEMENTATION.
  METHOD constructor.
    DATA: lv_seed TYPE string,
          lv_hash TYPE string.

    lv_seed = |{ sy-mandt }| && |{ sy-uname }| && |{ sy-repid }| && |{ sy-datum }| && |{ sy-uzeit }|. "#EC NOTEXT
    TRY.
        cl_abap_message_digest=>calculate_hash_for_char(
          EXPORTING
            if_algorithm  = 'MD5' "#EC NOTEXT
            if_data       = lv_seed
          IMPORTING
            ef_hashstring = lv_hash ).
        mv_session_id = lv_hash(32).
      CATCH cx_abap_message_digest.
        mv_session_id = |{ sy-uname }{ sy-uzeit }|. "#EC NOTEXT
    ENDTRY.
  ENDMETHOD.

  METHOD cleanup_expired.
    DATA: lv_date TYPE sy-datum,
          lv_time TYPE sy-uzeit,
          lv_secs TYPE i.

    lv_secs = gc_timeout_secs * -1.
    TRY.
        cl_abap_tstmp=>td_add(
          EXPORTING
            date     = sy-datum
            time     = sy-uzeit
            secs     = lv_secs
          IMPORTING
            res_date = lv_date
            res_time = lv_time ).
      CATCH cx_parameter_invalid.
        lv_date = sy-datum.
        lv_time = sy-uzeit.
    ENDTRY.

    DELETE FROM zflex_edit_lock
      WHERE tabname = @p_tabnam
        AND ( last_date < @lv_date
           OR ( last_date = @lv_date AND last_time < @lv_time ) ).
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
    ENDIF.
  ENDMETHOD.

  METHOD get_table_lock.
    cleanup_expired( ).
    SELECT SINGLE mandt, tabname, key_hash, uname, session_id,
                  fieldname, key_text, lock_date, lock_time,
                  last_date, last_time, progname
      FROM zflex_edit_lock
      INTO CORRESPONDING FIELDS OF @rs_lock
      WHERE tabname  = @p_tabnam
        AND key_hash = @gc_table_key_hash.
  ENDMETHOD.

  METHOD is_table_locked_by_other.
    DATA(ls_lock) = get_table_lock( ).
    rv_locked = abap_false.
    IF ls_lock-key_hash IS NOT INITIAL
       AND ls_lock-uname <> sy-uname.
      rv_locked = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD acquire_table_lock.
    DATA ls_lock TYPE zflex_edit_lock.

    rv_success = abap_false.
    cleanup_expired( ).

    SELECT SINGLE mandt, tabname, key_hash, uname, session_id,
                  fieldname, key_text, lock_date, lock_time,
                  last_date, last_time, progname
      FROM zflex_edit_lock
      INTO CORRESPONDING FIELDS OF @ls_lock
      WHERE tabname  = @p_tabnam
        AND key_hash = @gc_table_key_hash.
    IF sy-subrc = 0 AND ls_lock-uname <> sy-uname.
      RETURN.
    ENDIF.

    CLEAR ls_lock.
    ls_lock-mandt      = sy-mandt.
    ls_lock-tabname    = p_tabnam.
    ls_lock-key_hash   = gc_table_key_hash.
    ls_lock-uname      = sy-uname.
    ls_lock-session_id = mv_session_id.
    ls_lock-fieldname  = gc_table_field.
    ls_lock-key_text   = |TABLE={ p_tabnam }|. "#EC NOTEXT
    ls_lock-lock_date  = sy-datum.
    ls_lock-lock_time  = sy-uzeit.
    ls_lock-last_date  = sy-datum.
    ls_lock-last_time  = sy-uzeit.
    ls_lock-progname   = sy-repid.

    MODIFY zflex_edit_lock FROM @ls_lock.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
      rv_success = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD build_key_text.
    FIELD-SYMBOLS: <lv_key> TYPE any.
    DATA: lv_value TYPE string.

    CLEAR rv_key_text.
    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_fcat) WHERE key = 'X'. "#EC NOTEXT
      ASSIGN COMPONENT ls_fcat-fieldname OF STRUCTURE is_row TO <lv_key>.
      IF sy-subrc <> 0.
        CLEAR rv_key_text.
        RETURN.
      ENDIF.
      lv_value = |{ <lv_key> }|. "#EC NOTEXT
      rv_key_text = rv_key_text && |{ ls_fcat-fieldname }={ lv_value }| && '|'. "#EC NOTEXT
    ENDLOOP.
  ENDMETHOD.

  METHOD build_key_hash.
    DATA: lv_hash TYPE string.

    IF iv_key_text IS INITIAL.
      CLEAR rv_key_hash.
      RETURN.
    ENDIF.

    TRY.
        cl_abap_message_digest=>calculate_hash_for_char(
          EXPORTING
            if_algorithm  = 'MD5' "#EC NOTEXT
            if_data       = iv_key_text
          IMPORTING
            ef_hashstring = lv_hash ).
        rv_key_hash = lv_hash(32).
      CATCH cx_abap_message_digest.
        rv_key_hash = iv_key_text.
    ENDTRY.
  ENDMETHOD.

  METHOD release_session.
    DELETE FROM zflex_edit_lock
      WHERE tabname    = @p_tabnam
        AND uname      = @sy-uname
        AND session_id = @mv_session_id.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_audit_logger IMPLEMENTATION.
  METHOD generate_change_id.
    DATA: lv_uuid TYPE sysuuid_x16.
    TRY.
        lv_uuid = cl_system_uuid=>create_uuid_x16_static( ).
        rv_id = lv_uuid.
      CATCH cx_uuid_error.
        rv_id = |{ sy-uname }{ sy-datum }{ sy-uzeit }{ sy-tabix }|. "#EC NOTEXT
    ENDTRY.
  ENDMETHOD.

  METHOD get_timestamp.
    GET TIME STAMP FIELD rv_tstamp.
  ENDMETHOD.

  METHOD log_access.
    DATA: ls_hdr TYPE ty_audit_hdr.
    ls_hdr-change_id   = generate_change_id( ).
    ls_hdr-tabname     = iv_tabname.
    ls_hdr-action      = iv_action.
    ls_hdr-uname       = sy-uname.
    ls_hdr-change_date = sy-datum.
    ls_hdr-change_time = sy-uzeit.
    ls_hdr-tstamp      = get_timestamp( ).
    ls_hdr-row_count   = iv_rows.
    ls_hdr-source      = iv_source.
    ls_hdr-progname    = sy-repid.
    ls_hdr-status      = 'S'. "#EC NOTEXT
    ls_hdr-description = iv_desc.
    ls_hdr-client      = sy-mandt.
    INSERT zflex_audit_hdr FROM @ls_hdr.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
    ENDIF.
  ENDMETHOD.

  METHOD log_changes.
    DATA: ls_hdr      TYPE ty_audit_hdr,
          lt_items    TYPE TABLE OF ty_audit_itm,
          ls_item     TYPE ty_audit_itm,
          lv_item_no  TYPE n LENGTH 6,
          lv_key_text TYPE string,
          lv_key_hash TYPE c LENGTH 32,
          lv_old_key  TYPE string,
          lv_old_hash TYPE c LENGTH 32,
          lv_found    TYPE abap_bool.
    FIELD-SYMBOLS: <ls_new>     TYPE any,
                   <ls_old>     TYPE any,
                   <lv_new_val> TYPE any,
                   <lv_old_val> TYPE any.

    DATA(lv_change_id) = generate_change_id( ).

    ls_hdr-change_id   = lv_change_id.
    ls_hdr-tabname     = iv_tabname.
    ls_hdr-action      = iv_action.
    ls_hdr-uname       = sy-uname.
    ls_hdr-change_date = sy-datum.
    ls_hdr-change_time = sy-uzeit.
    ls_hdr-tstamp      = get_timestamp( ).
    ls_hdr-row_count   = lines( it_new_data ).
    ls_hdr-source      = iv_source.
    ls_hdr-progname    = sy-repid.
    ls_hdr-status      = 'S'. "#EC NOTEXT
    ls_hdr-description = iv_desc.

    lv_item_no = '000000'.

    LOOP AT it_new_data ASSIGNING <ls_new>.
      IF go_lock_mgr IS BOUND.
        lv_key_text = go_lock_mgr->build_key_text( <ls_new> ).
        lv_key_hash = go_lock_mgr->build_key_hash( lv_key_text ).
      ENDIF.

      IF iv_action = 'UPDATE' AND it_old_data IS SUPPLIED "#EC NOTEXT
         AND it_old_data IS NOT INITIAL.
        lv_found = abap_false.
        LOOP AT it_old_data ASSIGNING <ls_old>.
          lv_old_key  = go_lock_mgr->build_key_text( <ls_old> ).
          lv_old_hash = go_lock_mgr->build_key_hash( lv_old_key ).
          IF lv_old_hash = lv_key_hash AND lv_key_hash IS NOT INITIAL.
            lv_found = abap_true.
            EXIT.
          ENDIF.
        ENDLOOP.

        IF lv_found = abap_true.
          LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_fcat)
            WHERE key <> 'X'. "#EC NOTEXT
            IF ls_fcat-fieldname = 'IS_NEW_ROW' OR ls_fcat-fieldname = 'CELL_STYLES' "#EC NOTEXT
               OR ls_fcat-fieldname = 'ROW_STATUS' OR ls_fcat-fieldname = 'CELL_COLORS' "#EC NOTEXT
               OR ls_fcat-fieldname = 'LOCK_OWNER' OR ls_fcat-fieldname = 'LOCK_FIELD' "#EC NOTEXT
               OR ls_fcat-fieldname = 'LOCK_INFO' OR ls_fcat-fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
              CONTINUE.
            ENDIF.
            ASSIGN COMPONENT ls_fcat-fieldname OF STRUCTURE <ls_new> TO <lv_new_val>.
            IF sy-subrc <> 0. CONTINUE. ENDIF.
            ASSIGN COMPONENT ls_fcat-fieldname OF STRUCTURE <ls_old> TO <lv_old_val>.
            IF sy-subrc <> 0. CONTINUE. ENDIF.
            IF <lv_new_val> <> <lv_old_val>.
              lv_item_no = lv_item_no + 1.
              CLEAR ls_item.
              ls_item-change_id = lv_change_id.
              ls_item-item_no   = lv_item_no.
              ls_item-key_hash  = lv_key_hash.
              ls_item-key_text  = lv_key_text.
              ls_item-fieldname = ls_fcat-fieldname.
              ls_item-old_value = |{ <lv_old_val> }|. "#EC NOTEXT
              ls_item-new_value = |{ <lv_new_val> }|. "#EC NOTEXT
              APPEND ls_item TO lt_items.
            ENDIF.
          ENDLOOP.
        ELSE.
          lv_item_no = lv_item_no + 1.
          CLEAR ls_item.
          ls_item-change_id = lv_change_id.
          ls_item-item_no   = lv_item_no.
          ls_item-key_hash  = lv_key_hash.
          ls_item-key_text  = lv_key_text.
          APPEND ls_item TO lt_items.
        ENDIF.

      ELSE.
        lv_item_no = lv_item_no + 1.
        CLEAR ls_item.
        ls_item-change_id = lv_change_id.
        ls_item-item_no   = lv_item_no.
        ls_item-key_hash  = lv_key_hash.
        ls_item-key_text  = lv_key_text.
        APPEND ls_item TO lt_items.
      ENDIF.
    ENDLOOP.

    ls_hdr-client = sy-mandt.
    INSERT zflex_audit_hdr FROM @ls_hdr.
    IF lt_items IS NOT INITIAL.
      LOOP AT lt_items ASSIGNING FIELD-SYMBOL(<ls_itm>).
        <ls_itm>-client = sy-mandt.
      ENDLOOP.
      INSERT zflex_audit_itm FROM TABLE @lt_items.
    ENDIF.
    COMMIT WORK AND WAIT.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_event_handler IMPLEMENTATION.
  METHOD on_toolbar.
    DATA: ls_toolbar TYPE stb_button,
          lt_orig    TYPE ttb_button.
    lt_orig = e_object->mt_toolbar.
    CLEAR e_object->mt_toolbar.

    " --- Save ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZSAVE'. "#EC NOTEXT
    ls_toolbar-icon      = icon_system_save.
    ls_toolbar-text      = 'Save'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Save changes to Database'. "#EC NOTEXT
    IF gv_edit_mode = abap_false OR gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Create ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZCREATE'. "#EC NOTEXT
    ls_toolbar-icon      = icon_create.
    ls_toolbar-text      = 'Create'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Create new entry'. "#EC NOTEXT
    IF gv_edit_mode = abap_false OR gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Edit ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZEDIT'. "#EC NOTEXT
    ls_toolbar-icon      = icon_change.
    ls_toolbar-text      = 'Edit'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Toggle edit mode'. "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Delete ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZDELETE'. "#EC NOTEXT
    ls_toolbar-icon      = icon_delete.
    ls_toolbar-text      = 'Delete'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Delete selected entries'. "#EC NOTEXT
    IF gv_edit_mode = abap_false OR gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Separator ---
    CLEAR ls_toolbar.
    ls_toolbar-butn_type = 3.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Filter ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZFILTER'. "#EC NOTEXT
    ls_toolbar-icon      = icon_filter.
    ls_toolbar-text      = 'Filter'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Change selection criteria (SE16N style)'. "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Excel Sync ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZEXCEL'. "#EC NOTEXT
    ls_toolbar-icon      = icon_xls.
    ls_toolbar-text      = 'Excl Sync'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Excel Download / Upload'. "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Push to Cloud ---
    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZPUSH'. "#EC NOTEXT
    ls_toolbar-icon      = icon_trend_up.
    ls_toolbar-text      = 'Sync to Cloud'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Upload persisted DB data to Google Sheet'. "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    " --- Separator ---
    CLEAR ls_toolbar.
    ls_toolbar-butn_type = 3.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    IF gv_upload_preview = abap_true.
      CLEAR ls_toolbar.
      ls_toolbar-function  = 'ZUPACC'. "#EC NOTEXT
      ls_toolbar-text      = 'Accept Upload'. "#EC NOTEXT
      ls_toolbar-quickinfo = 'Save preview rows to Database'. "#EC NOTEXT
      IF gv_upload_has_errors = abap_true.
        ls_toolbar-disabled = 'X'. "#EC NOTEXT
      ENDIF.
      APPEND ls_toolbar TO e_object->mt_toolbar.

      CLEAR ls_toolbar.
      ls_toolbar-function  = 'ZUPCANCEL'. "#EC NOTEXT
      ls_toolbar-text      = 'Cancel Upload'. "#EC NOTEXT
      ls_toolbar-quickinfo = 'Discard upload preview and restore previous data'. "#EC NOTEXT
      APPEND ls_toolbar TO e_object->mt_toolbar.

      CLEAR ls_toolbar.
      ls_toolbar-butn_type = 3.
      APPEND ls_toolbar TO e_object->mt_toolbar.
    ENDIF.

    APPEND LINES OF lt_orig TO e_object->mt_toolbar.
  ENDMETHOD.

  METHOD on_user_command.
    CASE e_ucomm.
      WHEN 'ZSAVE'. "#EC NOTEXT
        go_alv_report->save_data( ).
      WHEN 'ZUPACC'. "#EC NOTEXT
        go_alv_report->accept_upload_preview( ).
      WHEN 'ZUPCANCEL'. "#EC NOTEXT
        go_alv_report->cancel_upload_preview( ).
      WHEN 'ZCREATE'. "#EC NOTEXT
        go_alv_report->create_entry( ).
      WHEN 'ZEDIT'. "#EC NOTEXT
        go_alv_report->toggle_edit( ).
      WHEN 'ZDELETE'. "#EC NOTEXT
        go_alv_report->delete_entries( ).
      WHEN 'ZEXCEL'. "#EC NOTEXT
        go_alv_report->excel_sync_popup( ).
      WHEN 'ZCLOUD'. "#EC NOTEXT
        go_alv_report->sync_from_cloud( ).
      WHEN 'ZPUSH'. "#EC NOTEXT
        go_alv_report->push_to_cloud( ).
      WHEN 'ZFILTER'. "#EC NOTEXT
        IF gv_upload_preview = abap_true.
          MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
          RETURN.
        ENDIF.
        " Re-filter: show dialog, re-select, refresh ALV
        DATA(lv_cancelled) = go_filter->show_dialog( ).
        IF lv_cancelled = abap_false.
          go_filter->execute_select( ).
          go_grid->refresh_table_display( ).
          go_alv_report->refresh_header( ).
          DATA(lv_selected_rows) = lines( <dyn_table> ).
          MESSAGE s030(zflex_msg) WITH lv_selected_rows.
        ENDIF.
    ENDCASE.
  ENDMETHOD.

  METHOD on_data_changed.
    DATA: lv_lock_updated TYPE abap_bool.
    FIELD-SYMBOLS: <lv_lock_owner_chg> TYPE any,
                   <lv_lock_field_chg> TYPE any,
                   <lv_lock_info_chg> TYPE any,
                   <lt_color_chg> TYPE lvc_t_scol.
    DATA: ls_color_chg TYPE lvc_s_scol.

    IF go_lock_mgr IS NOT BOUND OR er_data_changed IS NOT BOUND.
      RETURN.
    ENDIF.

    LOOP AT er_data_changed->mt_mod_cells INTO DATA(ls_mod_cell).
      IF ls_mod_cell-fieldname = 'LOCK_OWNER' "#EC NOTEXT
        OR ls_mod_cell-fieldname = 'LOCK_FIELD' "#EC NOTEXT
        OR ls_mod_cell-fieldname = 'LOCK_INFO' "#EC NOTEXT
        OR ls_mod_cell-fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
        CONTINUE.
      ENDIF.

      READ TABLE <dyn_table> ASSIGNING <dyn_wa> INDEX ls_mod_cell-row_id.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      gv_has_changes = abap_true.

      ASSIGN COMPONENT 'LOCK_OWNER' OF STRUCTURE <dyn_wa> TO <lv_lock_owner_chg>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_lock_owner_chg> = sy-uname.
      ENDIF.
      ASSIGN COMPONENT 'LOCK_FIELD' OF STRUCTURE <dyn_wa> TO <lv_lock_field_chg>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_lock_field_chg> = lcl_lock_manager=>gc_table_field.
      ENDIF.
      ASSIGN COMPONENT 'LOCK_INFO' OF STRUCTURE <dyn_wa> TO <lv_lock_info_chg>. "#EC NOTEXT
      IF sy-subrc = 0.
        <lv_lock_info_chg> = |Table locked by me; changed field { ls_mod_cell-fieldname }|. "#EC NOTEXT
      ENDIF.
      PERFORM set_row_status USING <dyn_wa> 'Edit'. "#EC NOTEXT
      ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE <dyn_wa> TO <lt_color_chg>. "#EC NOTEXT
      IF sy-subrc = 0.
        CLEAR <lt_color_chg>.
        CLEAR ls_color_chg.
        ls_color_chg-fname = ls_mod_cell-fieldname.
        ls_color_chg-color-col = 3.
        ls_color_chg-color-int = 1.
        ls_color_chg-color-inv = 0.
        APPEND ls_color_chg TO <lt_color_chg>.
      ENDIF.
      lv_lock_updated = abap_true.
    ENDLOOP.

    IF lv_lock_updated = abap_true.
      IF go_grid IS BOUND.
        go_grid->refresh_table_display( ).
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_gsheet_event_handler IMPLEMENTATION.
  METHOD on_toolbar.
    CLEAR e_object->mt_toolbar.

    DATA ls_toolbar TYPE stb_button.

    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZGS_REFRESH'. "#EC NOTEXT
    ls_toolbar-icon      = icon_refresh.
    ls_toolbar-text      = 'Refresh'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Reload Google Sheet and highlight differences'. "#EC NOTEXT
    APPEND ls_toolbar TO e_object->mt_toolbar.

    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZGS_EXPORT'. "#EC NOTEXT
    ls_toolbar-icon      = icon_trend_up.
    ls_toolbar-text      = 'Export to Cloud'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Create Google Sheet instance and upload current SAP DB data'. "#EC NOTEXT
    IF gv_gsheet_can_export = abap_false.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.

    CLEAR ls_toolbar.
    ls_toolbar-function  = 'ZCLOUD'. "#EC NOTEXT
    ls_toolbar-icon      = icon_trend_down.
    ls_toolbar-text      = 'Sync from Cloud'. "#EC NOTEXT
    ls_toolbar-quickinfo = 'Download Google Sheet rows into ALV preview without saving DB'. "#EC NOTEXT
    IF gv_edit_mode = abap_false
       OR gv_upload_preview = abap_true
       OR gv_gsheet_has_instance = abap_false.
      ls_toolbar-disabled = 'X'. "#EC NOTEXT
    ENDIF.
    APPEND ls_toolbar TO e_object->mt_toolbar.
  ENDMETHOD.

  METHOD on_user_command.
    CASE e_ucomm.
      WHEN 'ZGS_REFRESH'. "#EC NOTEXT
        PERFORM reload_gsheet_alv.
      WHEN 'ZGS_EXPORT'. "#EC NOTEXT
        go_alv_report->export_to_cloud( ).
      WHEN 'ZCLOUD'. "#EC NOTEXT
        go_alv_report->sync_from_cloud( ).
    ENDCASE.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_alv_report IMPLEMENTATION.
  METHOD display_alv.
    IF go_grid IS BOUND.
      refresh_header( ).
      IF gv_edit_mode = abap_true.
        go_grid->set_ready_for_input( i_ready_for_input = 1 ).
      ELSE.
        go_grid->set_ready_for_input( i_ready_for_input = 0 ).
      ENDIF.
      refresh_mode_title( ).
      go_grid->refresh_table_display( ).
      IF go_sheet_grid IS BOUND.
        go_sheet_grid->refresh_table_display( ).
      ENDIF.
      RETURN.
    ENDIF.

    IF gv_edit_mode = abap_true
       AND go_lock_mgr IS BOUND
       AND go_lock_mgr->acquire_table_lock( ) = abap_false.
      DATA(ls_init_table_lock) = go_lock_mgr->get_table_lock( ).
      gv_edit_mode = abap_false.
      go_dyn_handler->set_edit_mode( gv_edit_mode ).
      MESSAGE s031(zflex_msg) WITH p_tabnam ls_init_table_lock-uname DISPLAY LIKE 'E'.
    ELSEIF gv_edit_mode = abap_true.
      go_filter->execute_select( ).
    ENDIF.

    " We use the standard screen container cl_gui_container=>screen0.
    " This is 100% robust under WebGUI & WinGUI and avoids CNTL_ERROR.
    DATA(lo_parent) = cl_gui_container=>screen0.

    " Main screen: Google Sheet ALV on top, Z-table ALV below.
    CREATE OBJECT go_main_splitter
      EXPORTING
        parent  = lo_parent
        rows    = 2
        columns = 1.

    go_main_splitter->set_row_height( id = 1 height = 35 ).
    go_main_splitter->set_border( border = space ).

    DATA(lo_sheet_cont) = go_main_splitter->get_container( row = 1 column = 1 ).
    DATA(lo_ztab_cont) = go_main_splitter->get_container( row = 2 column = 1 ).

    " Splitter for SE16N-style header and main Z-table grid.
    CREATE OBJECT go_header_splitter
      EXPORTING
        parent  = lo_ztab_cont
        rows    = 2
        columns = 1.

    go_header_splitter->set_row_height( id = 1 height = 15 ).
    go_header_splitter->set_border( border = space ).

    go_top_cont = go_header_splitter->get_container( row = 1 column = 1 ).
    DATA(lo_bot_cont) = go_header_splitter->get_container( row = 2 column = 1 ).

    " Render SE16N header
    CREATE OBJECT go_hdr_doc.
    build_html_header( go_hdr_doc ).
    go_hdr_doc->display_document( parent = go_top_cont ).

    " Create ALV grid
    CREATE OBJECT go_grid
      EXPORTING
        i_parent = lo_bot_cont.

    " Register handlers BEFORE first display
    SET HANDLER lcl_event_handler=>on_toolbar FOR go_grid.
    SET HANDLER lcl_event_handler=>on_user_command FOR go_grid.
    SET HANDLER lcl_event_handler=>on_data_changed FOR go_grid.
    go_grid->register_edit_event( i_event_id = cl_gui_alv_grid=>mc_evt_enter ).
    go_grid->register_edit_event( i_event_id = cl_gui_alv_grid=>mc_evt_modified ).

    DATA: ls_layout TYPE lvc_s_layo.
    ls_layout-zebra      = 'X'. "#EC NOTEXT
    ls_layout-sel_mode   = 'A'. "#EC NOTEXT
    ls_layout-cwidth_opt = 'X'. "#EC NOTEXT
    ls_layout-stylefname = 'CELL_STYLES'. "#EC NOTEXT
    ls_layout-ctab_fname = 'CELL_COLORS'. "#EC NOTEXT
    ls_layout-excp_fname = 'ROW_STATUS'. "#EC NOTEXT
    ls_layout-excp_led   = 'X'. "#EC NOTEXT
    IF gv_edit_mode = abap_true.
      ls_layout-grid_title = |Editable mode - { p_tabnam }|. "#EC NOTEXT
    ELSE.
      ls_layout-grid_title = |Display mode (read-only) - { p_tabnam }|. "#EC NOTEXT
    ENDIF.

    DATA: lt_exclude TYPE ui_functions.
    APPEND cl_gui_alv_grid=>mc_fc_loc_cut TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_copy TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_paste TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_undo TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_append_row TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_insert_row TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_delete_row TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_copy_row TO lt_exclude.
    APPEND cl_gui_alv_grid=>mc_fc_loc_paste_new_row TO lt_exclude.

    go_grid->set_table_for_first_display(
      EXPORTING
        is_layout            = ls_layout
        it_toolbar_excluding = lt_exclude
      CHANGING
        it_outtab            = <dyn_table>
        it_fieldcatalog      = go_dyn_handler->mt_fieldcat
      EXCEPTIONS
        OTHERS               = 4 ).

    IF gv_edit_mode = abap_true.
      go_grid->set_ready_for_input( i_ready_for_input = 1 ).
    ELSE.
      go_grid->set_ready_for_input( i_ready_for_input = 0 ).
    ENDIF.
    refresh_mode_title( ).

    PERFORM load_gsheet_data.

    CREATE OBJECT go_sheet_grid
      EXPORTING
        i_parent = lo_sheet_cont.

    SET HANDLER lcl_gsheet_event_handler=>on_toolbar FOR go_sheet_grid.
    SET HANDLER lcl_gsheet_event_handler=>on_user_command FOR go_sheet_grid.

    DATA(ls_sheet_layout) = VALUE lvc_s_layo(
      zebra      = 'X' "#EC NOTEXT
      sel_mode   = 'A' "#EC NOTEXT
      cwidth_opt = 'X' "#EC NOTEXT
      ctab_fname = 'CELL_COLORS' "#EC NOTEXT
      grid_title = gv_gsheet_status ).

    DATA lt_sheet_fcat TYPE lvc_t_fcat.
    PERFORM build_gsheet_fieldcat CHANGING lt_sheet_fcat.

    go_sheet_grid->set_table_for_first_display(
      EXPORTING
        is_layout       = ls_sheet_layout
      CHANGING
        it_outtab       = <gsheet_table>
        it_fieldcatalog = lt_sheet_fcat
      EXCEPTIONS
        OTHERS          = 4 ).
    go_sheet_grid->set_toolbar_interactive( ).
  ENDMETHOD.

  METHOD build_html_header.
    DATA: lv_info TYPE sdydo_text_element.

    io_dd->add_text( text = 'General Table Display' sap_style = 'HEADING' ). "#EC NOTEXT
    io_dd->new_line( ).

    lv_info = |Table: { p_tabnam }|. "#EC NOTEXT
    io_dd->add_text( text = lv_info sap_style = 'KEY' ). "#EC NOTEXT
    io_dd->add_gap( width = 5 ).

    DATA(lv_count) = lines( <dyn_table> ).
    lv_info = |Entries: { lv_count }|. "#EC NOTEXT
    io_dd->add_text( text = lv_info ).
    io_dd->add_gap( width = 5 ).

    " Show filter status
    lv_info = |[ { go_filter->get_filter_text( ) } ]|. "#EC NOTEXT
    io_dd->add_text( text = lv_info sap_style = 'KEY' ). "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      io_dd->add_gap( width = 5 ).
      IF gv_upload_has_errors = abap_true.
        io_dd->add_text( text = '[ Upload preview has errors ]' sap_style = 'KEY' ). "#EC NOTEXT
      ELSE.
        io_dd->add_text( text = '[ Upload preview ready ]' sap_style = 'KEY' ). "#EC NOTEXT
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD refresh_mode_title.
    DATA: ls_layout TYPE lvc_s_layo.

    IF go_grid IS NOT BOUND.
      RETURN.
    ENDIF.

    ls_layout-zebra      = 'X'. "#EC NOTEXT
    ls_layout-sel_mode   = 'A'. "#EC NOTEXT
    ls_layout-cwidth_opt = 'X'. "#EC NOTEXT
    ls_layout-stylefname = 'CELL_STYLES'. "#EC NOTEXT
    ls_layout-ctab_fname = 'CELL_COLORS'. "#EC NOTEXT
    ls_layout-excp_fname = 'ROW_STATUS'. "#EC NOTEXT
    ls_layout-excp_led   = 'X'. "#EC NOTEXT
    IF gv_upload_preview = abap_true.
      IF gv_upload_has_errors = abap_true.
        ls_layout-grid_title = |Upload preview has errors - { p_tabnam }|. "#EC NOTEXT
      ELSE.
        ls_layout-grid_title = |Upload preview ready - { p_tabnam }|. "#EC NOTEXT
      ENDIF.
    ELSEIF gv_edit_mode = abap_true.
      ls_layout-grid_title = |Editable mode - { p_tabnam }|. "#EC NOTEXT
    ELSE.
      ls_layout-grid_title = |Display mode (read-only) - { p_tabnam }|. "#EC NOTEXT
    ENDIF.

    go_grid->set_frontend_layout( is_layout = ls_layout ).
  ENDMETHOD.

  METHOD refresh_header.
    IF go_top_cont IS NOT BOUND.
      RETURN.
    ENDIF.

    IF go_hdr_doc IS NOT BOUND.
      CREATE OBJECT go_hdr_doc.
    ELSE.
      go_hdr_doc->initialize_document( ).
    ENDIF.

    build_html_header( go_hdr_doc ).
    go_hdr_doc->display_document( parent = go_top_cont ).
    cl_gui_cfw=>flush( ).
  ENDMETHOD.

  METHOD save_data.
    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    go_grid->check_changed_data( ).

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s004(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_save_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s033(zflex_msg) WITH p_tabnam ls_save_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    LOOP AT <dyn_table> ASSIGNING FIELD-SYMBOL(<ls_save_check>).
      ASSIGN COMPONENT 'ROW_STATUS' OF STRUCTURE <ls_save_check> TO FIELD-SYMBOL(<lv_save_status>). "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_save_status> = '1'.
        MESSAGE s034(zflex_msg) DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
    ENDLOOP.

    DATA: dref_flat_tab TYPE REF TO data,
          dref_old_tab  TYPE REF TO data.
    FIELD-SYMBOLS: <flat_tab> TYPE STANDARD TABLE,
                   <flat_wa>  TYPE ANY,
                   <old_tab>  TYPE STANDARD TABLE.
    CREATE DATA dref_flat_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_flat_tab->* TO <flat_tab>.

    " Capture old data for audit delta comparison
    CREATE DATA dref_old_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_old_tab->* TO <old_tab>.
    READ TABLE go_filter->mt_where_clauses INTO DATA(ls_audit_where)
      WITH KEY tablename = p_tabnam.
    IF sy-subrc = 0 AND ls_audit_where-where_tab IS NOT INITIAL.
      SELECT * FROM (p_tabnam) INTO TABLE @<old_tab>
        WHERE (ls_audit_where-where_tab).
    ELSE.
      SELECT * FROM (p_tabnam) INTO TABLE @<old_tab>.
    ENDIF.

    LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
      DATA(lv_missing_key) = abap_false.
      DATA(lv_missing_key_field) = VALUE fieldname( ).
      PERFORM validate_required_keys
        USING <dyn_wa>
        CHANGING lv_missing_key lv_missing_key_field.
      IF lv_missing_key = abap_true.
        IF go_grid IS BOUND.
          go_grid->refresh_table_display( ).
          cl_gui_cfw=>flush( ).
        ENDIF.
        MESSAGE s035(zflex_msg) WITH lv_missing_key_field DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      APPEND INITIAL LINE TO <flat_tab> ASSIGNING <flat_wa>.
      MOVE-CORRESPONDING <dyn_wa> TO <flat_wa>.
    ENDLOOP.

    MODIFY (p_tabnam) FROM TABLE <flat_tab>.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
      MESSAGE s036(zflex_msg).

      " Reset IS_NEW_ROW flag since all data is now persisted
      FIELD-SYMBOLS: <lv_is_new> TYPE any.
      LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
        ASSIGN COMPONENT 'IS_NEW_ROW' OF STRUCTURE <dyn_wa> TO <lv_is_new>. "#EC NOTEXT
        IF sy-subrc = 0.
          CLEAR <lv_is_new>.
        ENDIF.
      ENDLOOP.

      IF go_lock_mgr IS BOUND.
        go_lock_mgr->release_session( ).
      ENDIF.
      PERFORM leave_edit_after_write.
      PERFORM lock_existing_keys USING <dyn_table>.
      LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
        PERFORM set_row_status USING <dyn_wa> 'Updated'. "#EC NOTEXT
        PERFORM recolor_row_cells USING <dyn_wa> 3 5.
      ENDLOOP.
      gv_has_changes = abap_false.
      " === AUDIT LOG: UPDATE/CREATE ===
      IF go_audit IS BOUND.
        go_audit->log_changes(
          iv_tabname  = p_tabnam
          iv_action   = 'UPDATE' "#EC NOTEXT
          iv_source   = 'ALV' "#EC NOTEXT
          it_new_data = <flat_tab>
          it_old_data = <old_tab>
          iv_desc     = |Saved { lines( <flat_tab> ) } rows| ). "#EC NOTEXT
      ENDIF.
      go_grid->refresh_table_display( ).
    ELSE.
      ROLLBACK WORK.
      LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
        PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
      ENDLOOP.
      IF go_grid IS BOUND.
        go_grid->refresh_table_display( ).
      ENDIF.
      MESSAGE s037(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD accept_upload_preview.
    IF gv_upload_preview = abap_false.
      MESSAGE s025(zflex_msg) DISPLAY LIKE 'W'.
      RETURN.
    ENDIF.

    IF gv_upload_has_errors = abap_true.
      MESSAGE s019(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_grid IS BOUND.
      go_grid->check_changed_data( ).
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s004(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_accept_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s032(zflex_msg) WITH p_tabnam ls_accept_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    DATA: dref_flat_tab TYPE REF TO data,
          dref_old_tab TYPE REF TO data.
    FIELD-SYMBOLS: <flat_tab> TYPE STANDARD TABLE,
                   <flat_wa> TYPE ANY,
                   <old_tab> TYPE STANDARD TABLE.

    CREATE DATA dref_flat_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_flat_tab->* TO <flat_tab>.

    CREATE DATA dref_old_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_old_tab->* TO <old_tab>.
    READ TABLE go_filter->mt_where_clauses INTO DATA(ls_audit_where)
      WITH KEY tablename = p_tabnam.
    IF sy-subrc = 0 AND ls_audit_where-where_tab IS NOT INITIAL.
      SELECT * FROM (p_tabnam) INTO TABLE @<old_tab>
        WHERE (ls_audit_where-where_tab).
    ELSE.
      SELECT * FROM (p_tabnam) INTO TABLE @<old_tab>.
    ENDIF.

    LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
      APPEND INITIAL LINE TO <flat_tab> ASSIGNING <flat_wa>.
      MOVE-CORRESPONDING <dyn_wa> TO <flat_wa>.
    ENDLOOP.

    MODIFY (p_tabnam) FROM TABLE <flat_tab>.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
      IF go_lock_mgr IS BOUND.
        go_lock_mgr->release_session( ).
      ENDIF.
      PERFORM leave_edit_after_write.
      PERFORM lock_existing_keys USING <dyn_table>.
      LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
        PERFORM set_row_status USING <dyn_wa> 'Updated'. "#EC NOTEXT
      ENDLOOP.
      gv_upload_preview = abap_false.
      gv_upload_has_errors = abap_false.
      gv_has_changes = abap_false.
      CLEAR gr_upload_backup.

      IF go_audit IS BOUND.
        go_audit->log_changes(
          iv_tabname  = p_tabnam
          iv_action   = 'UPLOAD' "#EC NOTEXT
          iv_source   = 'EXCEL_UP' "#EC NOTEXT
          it_new_data = <flat_tab>
          it_old_data = <old_tab>
          iv_desc     = |Accepted upload preview with { lines( <flat_tab> ) } rows| ). "#EC NOTEXT
      ENDIF.

      IF go_grid IS BOUND.
        go_grid->refresh_table_display( ).
        go_grid->set_toolbar_interactive( ).
      ENDIF.
      refresh_header( ).
      DATA(lv_uploaded_rows) = lines( <flat_tab> ).
      MESSAGE s021(zflex_msg) WITH lv_uploaded_rows.
    ELSE.
      ROLLBACK WORK.
      gv_upload_has_errors = abap_true.
      LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
        PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
      ENDLOOP.
      IF go_grid IS BOUND.
        go_grid->refresh_table_display( ).
      ENDIF.
      refresh_header( ).
      MESSAGE s023(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD cancel_upload_preview.
    FIELD-SYMBOLS: <backup_tab> TYPE STANDARD TABLE.

    IF gv_upload_preview = abap_false.
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND.
      go_lock_mgr->release_session( ).
    ENDIF.
    PERFORM leave_edit_after_write.

    IF gr_upload_backup IS BOUND.
      ASSIGN gr_upload_backup->* TO <backup_tab>.
      IF sy-subrc = 0.
        <dyn_table> = <backup_tab>.
      ENDIF.
    ELSEIF go_filter IS BOUND.
      go_filter->execute_select( ).
    ENDIF.

    CLEAR gr_upload_backup.
    gv_upload_preview = abap_false.
    gv_upload_has_errors = abap_false.
    gv_has_changes = abap_false.

    PERFORM lock_existing_keys USING <dyn_table>.
    IF go_grid IS BOUND.
      go_grid->refresh_table_display( ).
      go_grid->set_toolbar_interactive( ).
    ENDIF.
    refresh_header( ).
    MESSAGE s022(zflex_msg).
  ENDMETHOD.

  METHOD create_entry.
    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF gv_edit_mode = abap_false.
      MESSAGE s014(zflex_msg) WITH 'creating rows' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_create_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s051(zflex_msg) WITH p_tabnam ls_create_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    FIELD-SYMBOLS: <lv_is_new> TYPE any.
    go_grid->check_changed_data( ).
    APPEND INITIAL LINE TO <dyn_table> ASSIGNING <dyn_wa>.
    ASSIGN COMPONENT 'IS_NEW_ROW' OF STRUCTURE <dyn_wa> TO <lv_is_new>. "#EC NOTEXT
    IF sy-subrc = 0.
      <lv_is_new> = abap_true.
    ENDIF.
    gv_has_changes = abap_true.
    go_grid->refresh_table_display( ).
    MESSAGE s050(zflex_msg).
  ENDMETHOD.

  METHOD delete_entries.
    DATA: lt_rows     TYPE lvc_t_row,
          lv_question TYPE c LENGTH 200,
          lv_deleted  TYPE i.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF gv_edit_mode = abap_false.
      MESSAGE s014(zflex_msg) WITH 'deleting rows' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    go_grid->check_changed_data( ).

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s004(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_delete_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s052(zflex_msg) WITH p_tabnam ls_delete_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    go_grid->get_selected_rows( IMPORTING et_index_rows = lt_rows ).
    IF lt_rows IS INITIAL.
      MESSAGE s053(zflex_msg) DISPLAY LIKE 'W'.
      RETURN.
    ENDIF.

    DATA(lv_answer) = ''.
    lv_question = |Delete { lines( lt_rows ) } selected row(s)?|. "#EC NOTEXT
    CALL FUNCTION 'POPUP_TO_CONFIRM' "#EC NOTEXT
      EXPORTING
        titlebar              = 'Confirm Delete' "#EC NOTEXT
        text_question         = lv_question
        text_button_1         = 'Yes' "#EC NOTEXT
        text_button_2         = 'No' "#EC NOTEXT
        default_button        = '2'
        display_cancel_button = ' '
      IMPORTING
        answer                = lv_answer.
    IF lv_answer <> '1'.
      RETURN.
    ENDIF.
    DATA: dref_del_tab TYPE REF TO data.
    FIELD-SYMBOLS: <del_tab> TYPE STANDARD TABLE,
                   <del_wa>  TYPE ANY.
    CREATE DATA dref_del_tab TYPE TABLE OF (p_tabnam).
    ASSIGN dref_del_tab->* TO <del_tab>.
    SORT lt_rows BY index DESCENDING.
    LOOP AT lt_rows INTO DATA(ls_row).
      READ TABLE <dyn_table> ASSIGNING <dyn_wa> INDEX ls_row-index.
      IF sy-subrc = 0.
        APPEND INITIAL LINE TO <del_tab> ASSIGNING <del_wa>.
        MOVE-CORRESPONDING <dyn_wa> TO <del_wa>.
        DELETE <dyn_table> INDEX ls_row-index.
      ENDIF.
    ENDLOOP.
    LOOP AT <del_tab> ASSIGNING <del_wa>.
      DELETE (p_tabnam) FROM <del_wa>.
      IF sy-subrc = 0.
        lv_deleted = lv_deleted + 1.
      ENDIF.
    ENDLOOP.
    IF lv_deleted <> lines( <del_tab> ).
      ROLLBACK WORK.
      LOOP AT lt_rows INTO DATA(ls_failed_row).
        READ TABLE <dyn_table> ASSIGNING <dyn_wa> INDEX ls_failed_row-index.
        IF sy-subrc = 0.
          PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
        ENDIF.
      ENDLOOP.
      IF go_grid IS BOUND.
        go_grid->refresh_table_display( ).
      ENDIF.
      MESSAGE s054(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.
    COMMIT WORK AND WAIT.
    IF go_lock_mgr IS BOUND.
      go_lock_mgr->release_session( ).
    ENDIF.
    PERFORM leave_edit_after_write.
    PERFORM lock_existing_keys USING <dyn_table>.
    gv_has_changes = abap_false.
    " === AUDIT LOG: DELETE ===
    IF go_audit IS BOUND.
      go_audit->log_changes(
        iv_tabname  = p_tabnam
        iv_action   = 'DELETE' "#EC NOTEXT
        iv_source   = 'ALV' "#EC NOTEXT
        it_new_data = <del_tab>
        iv_desc     = |Deleted { lv_deleted } rows| ). "#EC NOTEXT
    ENDIF.
    go_grid->refresh_table_display( ).
    MESSAGE s055(zflex_msg) WITH lv_deleted.
  ENDMETHOD.

  METHOD sync_from_cloud.
    DATA: lt_key_fcat      TYPE lvc_t_fcat,
          lt_seen_keys     TYPE HASHED TABLE OF zflex_edit_lock-key_hash WITH UNIQUE KEY table_line,
          lv_add_count     TYPE i,
          lv_update_count  TYPE i,
          lv_error_count   TYPE i,
          lv_diff_count    TYPE i,
          lv_missing_key   TYPE abap_bool,
          lv_missing_field TYPE fieldname,
          lv_key_text      TYPE string,
          lv_key_hash      TYPE zflex_edit_lock-key_hash,
          lv_error_text    TYPE string.

    FIELD-SYMBOLS: <ls_cloud>  TYPE any,
                   <ls_target> TYPE any,
                   <lv_cloud>  TYPE any,
                   <lv_target> TYPE any,
                   <lv_tech>   TYPE any.

    IF gv_edit_mode = abap_false.
      MESSAGE s014(zflex_msg) WITH 'syncing from Cloud' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s004(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->acquire_table_lock( ) = abap_false.
      DATA(ls_sync_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s070(zflex_msg) WITH p_tabnam ls_sync_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_grid IS BOUND.
      go_grid->check_changed_data( ).
    ENDIF.

    PERFORM load_gsheet_data.
    IF <gsheet_table> IS NOT ASSIGNED OR <gsheet_table> IS INITIAL.
      MESSAGE gv_gsheet_status TYPE 'S' DISPLAY LIKE 'W'. "#EC NOTEXT
      RETURN.
    ENDIF.

    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_key_fcat)
      WHERE key = 'X'. "#EC NOTEXT
      DATA(lv_is_key_tech) = abap_false.
      PERFORM is_cloud_compare_skip_field USING ls_key_fcat-fieldname CHANGING lv_is_key_tech.
      IF lv_is_key_tech = abap_false.
        APPEND ls_key_fcat TO lt_key_fcat.
      ENDIF.
    ENDLOOP.

    LOOP AT <gsheet_table> ASSIGNING <ls_cloud>.
      ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE <ls_cloud> TO FIELD-SYMBOL(<lv_cloud_message>). "#EC NOTEXT
      IF sy-subrc = 0.
        DATA(lv_cloud_message_text) = |{ <lv_cloud_message> }|. "#EC NOTEXT
        IF lv_cloud_message_text CP 'Invalid Cloud value*'. "#EC NOTEXT
          lv_error_count = lv_error_count + 1.
          CONTINUE.
        ENDIF.
      ENDIF.

      ASSIGN COMPONENT 'MANDT' OF STRUCTURE <ls_cloud> TO <lv_cloud>. "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_cloud> IS INITIAL.
        <lv_cloud> = sy-mandt.
      ENDIF.
      ASSIGN COMPONENT 'CLIENT' OF STRUCTURE <ls_cloud> TO <lv_cloud>. "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_cloud> IS INITIAL.
        <lv_cloud> = sy-mandt.
      ENDIF.

      CLEAR: lv_missing_key, lv_missing_field.
      PERFORM validate_required_keys
        USING <ls_cloud>
        CHANGING lv_missing_key lv_missing_field.
      IF lv_missing_key = abap_true.
        lv_error_count = lv_error_count + 1.
        CONTINUE.
      ENDIF.

      PERFORM build_cloud_compare_key
        USING <ls_cloud> lt_key_fcat sy-tabix
        CHANGING lv_key_text.
      lv_key_hash = go_lock_mgr->build_key_hash( lv_key_text ).
      IF lv_key_hash IS INITIAL.
        lv_error_count = lv_error_count + 1.
        PERFORM set_gsheet_message USING <ls_cloud> 'Cannot build key for Cloud row'. "#EC NOTEXT
        CONTINUE.
      ENDIF.

      READ TABLE lt_seen_keys TRANSPORTING NO FIELDS WITH TABLE KEY table_line = lv_key_hash.
      IF sy-subrc = 0.
        lv_error_count = lv_error_count + 1.
        PERFORM color_gsheet_row USING <ls_cloud> 'Duplicate key in Cloud data' 6. "#EC NOTEXT
        CONTINUE.
      ENDIF.
      INSERT lv_key_hash INTO TABLE lt_seen_keys.

      UNASSIGN <ls_target>.
      IF lt_key_fcat IS INITIAL.
        READ TABLE <dyn_table> INDEX sy-tabix ASSIGNING <ls_target>.
      ELSE.
        LOOP AT <dyn_table> ASSIGNING <dyn_wa>.
          DATA(lv_same_key) = abap_true.
          LOOP AT lt_key_fcat INTO DATA(ls_key_cmp).
            DATA(lv_cloud_key) = ||.
            DATA(lv_target_key) = ||.
            PERFORM get_compare_value USING <ls_cloud> ls_key_cmp CHANGING lv_cloud_key.
            PERFORM get_compare_value USING <dyn_wa>   ls_key_cmp CHANGING lv_target_key.
            IF lv_cloud_key <> lv_target_key.
              lv_same_key = abap_false.
              EXIT.
            ENDIF.
          ENDLOOP.
          IF lv_same_key = abap_true.
            ASSIGN <dyn_wa> TO <ls_target>.
            EXIT.
          ENDIF.
        ENDLOOP.
      ENDIF.

      IF <ls_target> IS ASSIGNED.
        CLEAR lv_diff_count.
        LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_cmp_fcat).
          DATA(lv_is_tech) = abap_false.
          PERFORM is_cloud_compare_skip_field USING ls_cmp_fcat-fieldname CHANGING lv_is_tech.
          IF lv_is_tech = abap_true.
            CONTINUE.
          ENDIF.

          DATA(lv_cloud_value) = ||.
          DATA(lv_target_value) = ||.
          PERFORM get_compare_value USING <ls_cloud>  ls_cmp_fcat CHANGING lv_cloud_value.
          PERFORM get_compare_value USING <ls_target> ls_cmp_fcat CHANGING lv_target_value.

          ASSIGN COMPONENT ls_cmp_fcat-fieldname OF STRUCTURE <ls_cloud> TO <lv_cloud>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.
          ASSIGN COMPONENT ls_cmp_fcat-fieldname OF STRUCTURE <ls_target> TO <lv_target>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          DATA(lv_cloud_input) = |{ <lv_cloud> }|. "#EC NOTEXT
          DATA(lv_cloud_norm) = ||.
          DATA(lv_target_raw) = |{ <lv_target> }|. "#EC NOTEXT
          PERFORM normalize_cloud_value
            USING ls_cmp_fcat lv_cloud_input
            CHANGING lv_cloud_norm.

          IF lv_cloud_value <> lv_target_value OR lv_cloud_norm <> lv_target_raw.
            TRY.
                <lv_target> = lv_cloud_norm.
                lv_diff_count = lv_diff_count + 1.
                PERFORM color_gsheet_cell USING <ls_target> ls_cmp_fcat-fieldname 3.
              CATCH cx_root INTO DATA(lx_assign).
                lv_error_count = lv_error_count + 1.
                lv_error_text = |Invalid Cloud value for { ls_cmp_fcat-fieldname }: { lx_assign->get_text( ) }|. "#EC NOTEXT
                PERFORM set_row_status USING <ls_target> 'Failed'. "#EC NOTEXT
                PERFORM set_row_message USING <ls_target> lv_error_text.
                PERFORM color_gsheet_cell USING <ls_target> ls_cmp_fcat-fieldname 6.
            ENDTRY.
          ENDIF.
        ENDLOOP.

        IF lv_diff_count > 0.
          lv_update_count = lv_update_count + 1.
          DATA(lv_cloud_message) = |Cloud sync preview: { lv_diff_count } field(s) updated|. "#EC NOTEXT
          PERFORM mark_cloud_preview_row USING <ls_target> lv_cloud_message.
        ENDIF.
      ELSE.
        APPEND INITIAL LINE TO <dyn_table> ASSIGNING <dyn_wa>.
        DATA(lv_new_row_error) = abap_false.
        LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_new_fcat).
          DATA(lv_new_is_tech) = abap_false.
          PERFORM is_cloud_compare_skip_field USING ls_new_fcat-fieldname CHANGING lv_new_is_tech.
          IF lv_new_is_tech = abap_true.
            CONTINUE.
          ENDIF.

          ASSIGN COMPONENT ls_new_fcat-fieldname OF STRUCTURE <ls_cloud> TO <lv_cloud>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.
          ASSIGN COMPONENT ls_new_fcat-fieldname OF STRUCTURE <dyn_wa> TO <lv_target>.
          IF sy-subrc <> 0.
            CONTINUE.
          ENDIF.

          TRY.
              DATA(lv_new_input) = |{ <lv_cloud> }|. "#EC NOTEXT
              DATA(lv_new_norm) = ||.
              PERFORM normalize_cloud_value
                USING ls_new_fcat lv_new_input
                CHANGING lv_new_norm.
              <lv_target> = lv_new_norm.
            CATCH cx_root INTO DATA(lx_new_assign).
              lv_new_row_error = abap_true.
              lv_error_count = lv_error_count + 1.
              lv_error_text = |Invalid Cloud value for { ls_new_fcat-fieldname }: { lx_new_assign->get_text( ) }|. "#EC NOTEXT
              PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
              PERFORM set_row_message USING <dyn_wa> lv_error_text.
              PERFORM color_gsheet_cell USING <dyn_wa> ls_new_fcat-fieldname 6.
          ENDTRY.
        ENDLOOP.
        PERFORM fill_generated_technical_keys USING <dyn_wa>.
        IF lv_new_row_error = abap_false.
          PERFORM mark_cloud_preview_row USING <dyn_wa> 'New row from Cloud preview'. "#EC NOTEXT
        ENDIF.
        ASSIGN COMPONENT 'IS_NEW_ROW' OF STRUCTURE <dyn_wa> TO <lv_tech>. "#EC NOTEXT
        IF sy-subrc = 0.
          <lv_tech> = abap_true.
        ENDIF.

        CLEAR: lv_missing_key, lv_missing_field.
        PERFORM validate_required_keys
          USING <dyn_wa>
          CHANGING lv_missing_key lv_missing_field.
        IF lv_missing_key = abap_true.
          lv_error_count = lv_error_count + 1.
          PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
        ELSEIF lv_new_row_error = abap_true.
          PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
        ELSE.
          lv_add_count = lv_add_count + 1.
          PERFORM set_row_status USING <dyn_wa> 'Edit'. "#EC NOTEXT
          PERFORM color_gsheet_row USING <dyn_wa> 'New row from Cloud preview' 3. "#EC NOTEXT
        ENDIF.
      ENDIF.
    ENDLOOP.

    IF lv_add_count > 0 OR lv_update_count > 0.
      gv_has_changes = abap_true.
    ENDIF.

    IF go_grid IS BOUND.
      go_grid->refresh_table_display( ).
      go_grid->set_toolbar_interactive( ).
    ENDIF.
    IF go_sheet_grid IS BOUND.
      go_sheet_grid->refresh_table_display( ).
    ENDIF.
    refresh_header( ).

    MESSAGE s071(zflex_msg) WITH lv_add_count lv_update_count lv_error_count.
  ENDMETHOD.

  METHOD push_to_cloud.
    DATA: lr_push_table TYPE REF TO data,
          lv_payload    TYPE string,
          lv_range      TYPE string,
          lv_row_count  TYPE i,
          lv_response   TYPE string.

    FIELD-SYMBOLS <lt_push_table> TYPE STANDARD TABLE.

    CLEAR gv_action_error.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_grid IS BOUND.
      go_grid->check_changed_data( ).
    ENDIF.

    IF gv_has_changes = abap_true.
      MESSAGE s038(zflex_msg) WITH 'pushing to Cloud' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_push_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s032(zflex_msg) WITH p_tabnam ls_push_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '03' "#EC NOTEXT
             ID 'TABLE' FIELD p_tabnam. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s003(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    CREATE DATA lr_push_table TYPE TABLE OF (p_tabnam).
    ASSIGN lr_push_table->* TO <lt_push_table>.
    IF sy-subrc <> 0 OR <lt_push_table> IS NOT ASSIGNED.
      MESSAGE s072(zflex_msg) WITH p_tabnam DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    PERFORM read_db_for_push CHANGING <lt_push_table>.
    IF gv_action_error IS NOT INITIAL.
      MESSAGE gv_action_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.
    PERFORM build_sheet_payload
      USING <lt_push_table>
      CHANGING lv_payload lv_range lv_row_count.
    IF gv_action_error IS NOT INITIAL.
      MESSAGE gv_action_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.
    PERFORM call_cpi_push
      USING lv_payload lv_range
      CHANGING lv_response.
    IF gv_action_error IS NOT INITIAL.
      MESSAGE gv_action_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    PERFORM reload_gsheet_alv.
    MESSAGE s073(zflex_msg) WITH lv_row_count.
  ENDMETHOD.

  METHOD export_to_cloud.
    DATA: lv_sheet_id  TYPE string,
          lv_worksheet TYPE string,
          lv_response  TYPE string.

    CLEAR gv_action_error.

    IF gv_gsheet_has_instance = abap_true.
      MESSAGE s074(zflex_msg) WITH p_tabnam.
      RETURN.
    ENDIF.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF go_grid IS BOUND.
      go_grid->check_changed_data( ).
    ENDIF.

    IF gv_has_changes = abap_true.
      MESSAGE s038(zflex_msg) WITH 'exporting to Cloud' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    IF go_lock_mgr IS BOUND
       AND go_lock_mgr->is_table_locked_by_other( ) = abap_true.
      DATA(ls_export_table_lock) = go_lock_mgr->get_table_lock( ).
      MESSAGE s075(zflex_msg) WITH p_tabnam ls_export_table_lock-uname DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    AUTHORITY-CHECK OBJECT 'S_TABU_NAM' "#EC NOTEXT
             ID 'ACTVT' FIELD '02' "#EC NOTEXT
             ID 'TABLE' FIELD 'ZTGSHEET_MAP'. "#EC NOTEXT
    IF sy-subrc <> 0.
      MESSAGE s076(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    PERFORM call_cpi_create_sheet
      CHANGING lv_sheet_id lv_worksheet lv_response.
    IF gv_action_error IS NOT INITIAL.
      MESSAGE gv_action_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    PERFORM upsert_gsheet_mapping
      USING lv_sheet_id lv_worksheet.
    IF gv_action_error IS NOT INITIAL.
      MESSAGE gv_action_error TYPE 'S' DISPLAY LIKE 'E'. "#EC NOTEXT
      RETURN.
    ENDIF.

    gv_gsheet_has_instance = abap_true.
    gv_gsheet_can_export = abap_false.

    MESSAGE s077(zflex_msg) WITH lv_sheet_id.
    me->push_to_cloud( ).
  ENDMETHOD.

  METHOD toggle_edit.
    DATA: lv_answer TYPE c LENGTH 1,
          lv_leave_edit TYPE abap_bool.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF gv_edit_mode = abap_true.
      IF go_grid IS BOUND.
        go_grid->check_changed_data( ).
      ENDIF.

      IF gv_has_changes = abap_true.
        CALL FUNCTION 'POPUP_TO_CONFIRM' "#EC NOTEXT
          EXPORTING
            titlebar              = 'Leave Edit Mode' "#EC NOTEXT
            text_question         = 'Leave edit mode? Save changes before switching to display mode?' "#EC NOTEXT
            text_button_1         = 'Save' "#EC NOTEXT
            text_button_2         = 'Discard' "#EC NOTEXT
            default_button        = '1'
            display_cancel_button = 'X' "#EC NOTEXT
          IMPORTING
            answer                = lv_answer.

        CASE lv_answer.
          WHEN '1'.
            save_data( ).
            lv_leave_edit = abap_true.
          WHEN '2'.
            IF go_lock_mgr IS BOUND.
              go_lock_mgr->release_session( ).
            ENDIF.
            gv_has_changes = abap_false.
            lv_leave_edit = abap_true.
          WHEN OTHERS.
            RETURN.
        ENDCASE.
      ELSE.
        IF go_lock_mgr IS BOUND.
          go_lock_mgr->release_session( ).
        ENDIF.
        lv_leave_edit = abap_true.
      ENDIF.

      IF lv_leave_edit = abap_true.
        gv_edit_mode = abap_false.
        go_dyn_handler->set_edit_mode( gv_edit_mode ).
        go_filter->execute_select( ).
      ENDIF.
    ELSE.
      IF go_lock_mgr IS BOUND
         AND go_lock_mgr->acquire_table_lock( ) = abap_false.
        DATA(ls_table_lock) = go_lock_mgr->get_table_lock( ).
        MESSAGE s032(zflex_msg) WITH p_tabnam ls_table_lock-uname DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      go_filter->execute_select( ).
      gv_edit_mode = abap_true.
      go_dyn_handler->set_edit_mode( gv_edit_mode ).
    ENDIF.

    PERFORM lock_existing_keys USING <dyn_table>.
    go_grid->set_frontend_fieldcatalog( it_fieldcatalog = go_dyn_handler->mt_fieldcat ).
    IF gv_edit_mode = abap_true.
      go_grid->set_ready_for_input( i_ready_for_input = 1 ).
    ELSE.
      go_grid->set_ready_for_input( i_ready_for_input = 0 ).
    ENDIF.
    refresh_mode_title( ).
    go_grid->refresh_table_display( ).
    go_grid->set_toolbar_interactive( ).
    cl_gui_cfw=>flush( ).
    refresh_header( ).
    cl_gui_cfw=>flush( ).

    IF gv_edit_mode = abap_true.
      MESSAGE s039(zflex_msg).
    ELSE.
      MESSAGE s049(zflex_msg).
    ENDIF.
  ENDMETHOD.

  METHOD excel_sync_popup.
    DATA: lv_answer   TYPE c LENGTH 1,
          lv_titlebar TYPE c LENGTH 80.

    IF gv_upload_preview = abap_true.
      MESSAGE s026(zflex_msg) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    lv_titlebar = |Table Action - [{ p_tabnam }]|. "#EC NOTEXT
    CALL FUNCTION 'POPUP_TO_CONFIRM' "#EC NOTEXT
      EXPORTING
        titlebar              = lv_titlebar
        text_question         = 'Choose an action:' "#EC NOTEXT
        text_button_1         = 'Download' "#EC NOTEXT
        icon_button_1         = 'ICON_EXPORT' "#EC NOTEXT
        text_button_2         = 'Upload' "#EC NOTEXT
        icon_button_2         = 'ICON_IMPORT' "#EC NOTEXT
        default_button        = '1'
        display_cancel_button = 'X' "#EC NOTEXT
      IMPORTING
        answer                = lv_answer.
    CASE lv_answer.
      WHEN '1'.
        go_excel_sync->download_excel( ).
      WHEN '2'.
        go_excel_sync->upload_excel( ).
        go_grid->refresh_table_display( ).
    ENDCASE.
  ENDMETHOD.
ENDCLASS.

" =====================================================================
" EVENT BLOCKS
" =====================================================================

INITIALIZATION.
  CREATE OBJECT go_sel_screen.
  gv_dummy_text = 'Z-Table Flex Manager Active'. "#EC NOTEXT

AT SELECTION-SCREEN.
  IF sy-dynnr = '1000' AND sy-ucomm = 'ONLI'. "#EC NOTEXT
    go_sel_screen->validate_input( ).
  ENDIF.

  IF sy-dynnr = '2000'.
    IF sy-ucomm = 'BACK' OR sy-ucomm = 'CANC' OR sy-ucomm = 'EXLI'. "#EC NOTEXT
      IF gv_upload_preview = abap_true.
        DATA(lv_exit_preview_answer) = ''.
        CALL FUNCTION 'POPUP_TO_CONFIRM' "#EC NOTEXT
          EXPORTING
            titlebar              = 'Cancel Upload Preview' "#EC NOTEXT
            text_question         = 'Cancel upload preview and leave this screen?' "#EC NOTEXT
            text_button_1         = 'Yes' "#EC NOTEXT
            text_button_2         = 'No' "#EC NOTEXT
            default_button        = '2'
            display_cancel_button = 'X' "#EC NOTEXT
          IMPORTING
            answer                = lv_exit_preview_answer.
        IF lv_exit_preview_answer <> '1'.
          RETURN.
        ENDIF.
        go_alv_report->cancel_upload_preview( ).
      ENDIF.
      IF go_lock_mgr IS BOUND.
        go_lock_mgr->release_session( ).
      ENDIF.
      IF go_grid IS BOUND.
        go_grid->free( ).
        CLEAR go_grid.
      ENDIF.
      IF go_sheet_grid IS BOUND.
        go_sheet_grid->free( ).
        CLEAR go_sheet_grid.
      ENDIF.
      CLEAR: go_top_cont,
             go_hdr_doc,
             go_main_splitter,
             go_header_splitter,
             gr_gsheet_table,
             gv_gsheet_status.
      UNASSIGN <gsheet_table>.
    ENDIF.
  ENDIF.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_tabnam.
  TYPES: BEGIN OF ty_f4_table,
           tabname TYPE dd02l-tabname,
           ddtext  TYPE dd02t-ddtext,
         END OF ty_f4_table.
  DATA: lt_f4 TYPE TABLE OF ty_f4_table.

  SELECT l~tabname, t~ddtext
    FROM dd02l AS l
    LEFT OUTER JOIN dd02t AS t ON t~tabname = l~tabname AND t~ddlanguage = @sy-langu AND t~as4local = 'A' "#EC NOTEXT
    INTO TABLE @lt_f4
    WHERE ( l~tabname LIKE 'Z%' OR l~tabname LIKE 'Y%' ) "#EC NOTEXT
      AND l~tabclass = 'TRANSP' "#EC NOTEXT
      AND l~as4local = 'A'. "#EC NOTEXT

  CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST' "#EC NOTEXT
    EXPORTING
      retfield    = 'TABNAME' "#EC NOTEXT
      dynpprog    = sy-repid
      dynpnr      = sy-dynnr
      dynprofield = 'P_TABNAM' "#EC NOTEXT
      value_org   = 'S' "#EC NOTEXT
    TABLES
      value_tab   = lt_f4
    EXCEPTIONS
      OTHERS      = 1.

AT SELECTION-SCREEN OUTPUT.
  IF sy-dynnr = '2000'.
    go_alv_report->display_alv( ).
  ENDIF.

FORM set_row_status USING ps_row TYPE any
                          pv_status TYPE string.
  lcl_alv_style=>set_row_status(
    EXPORTING iv_status = pv_status
    CHANGING  cs_row    = ps_row ).
ENDFORM.

FORM set_row_message USING ps_row TYPE any
                           pv_message TYPE string.
  lcl_alv_style=>set_row_message(
    EXPORTING iv_message = pv_message
    CHANGING  cs_row     = ps_row ).
ENDFORM.

FORM leave_edit_after_write.
  gv_edit_mode = abap_false.
  go_dyn_handler->set_edit_mode( gv_edit_mode ).

  IF go_grid IS BOUND.
    go_grid->set_frontend_fieldcatalog( it_fieldcatalog = go_dyn_handler->mt_fieldcat ).
    go_grid->set_ready_for_input( i_ready_for_input = 0 ).
    go_grid->set_toolbar_interactive( ).
  ENDIF.

  IF go_alv_report IS BOUND.
    go_alv_report->refresh_mode_title( ).
  ENDIF.
ENDFORM.

FORM validate_required_keys USING ps_row TYPE any
                            CHANGING pv_has_error TYPE abap_bool
                                     pv_fieldname TYPE fieldname.
  FIELD-SYMBOLS: <lv_key> TYPE any,
                 <lt_color> TYPE lvc_t_scol.
  DATA: ls_color TYPE lvc_s_scol,
        lv_message TYPE string.

  pv_has_error = abap_false.
  CLEAR pv_fieldname.

  ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE ps_row TO <lt_color>. "#EC NOTEXT
  IF sy-subrc = 0.
    CLEAR <lt_color>.
  ENDIF.

  LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_key_fcat)
    WHERE key = 'X'. "#EC NOTEXT
    DATA(lv_skip_required_key) = abap_false.
    PERFORM is_cloud_compare_skip_field USING ls_key_fcat-fieldname CHANGING lv_skip_required_key.
    IF lv_skip_required_key = abap_true
       AND ls_key_fcat-fieldname <> 'MANDT' "#EC NOTEXT
       AND ls_key_fcat-fieldname <> 'CLIENT'. "#EC NOTEXT
      CONTINUE.
    ENDIF.

    ASSIGN COMPONENT ls_key_fcat-fieldname OF STRUCTURE ps_row TO <lv_key>.
    IF sy-subrc <> 0.
      CONTINUE.
    ENDIF.

    IF ( ls_key_fcat-fieldname = 'MANDT' "#EC NOTEXT
         OR ls_key_fcat-fieldname = 'CLIENT' ) "#EC NOTEXT
       AND <lv_key> IS INITIAL.
      <lv_key> = sy-mandt.
      CONTINUE.
    ENDIF.

    IF <lv_key> IS INITIAL.
      pv_has_error = abap_true.
      pv_fieldname = ls_key_fcat-fieldname.
      PERFORM set_row_status USING ps_row 'Failed'. "#EC NOTEXT
      lv_message = |Missing key field { ls_key_fcat-fieldname }|. "#EC NOTEXT
      PERFORM set_row_message USING ps_row lv_message.

      IF <lt_color> IS ASSIGNED.
        CLEAR ls_color.
        ls_color-fname = ls_key_fcat-fieldname.
        ls_color-color-col = 6.
        ls_color-color-int = 1.
        ls_color-color-inv = 0.
        APPEND ls_color TO <lt_color>.
      ENDIF.
      RETURN.
    ENDIF.
  ENDLOOP.
ENDFORM.

FORM lock_existing_keys USING pt_dyn_table TYPE ANY TABLE.
  FIELD-SYMBOLS: <dyn_wa> TYPE ANY,
                 <lv_is_new> TYPE ANY,
                 <lv_row_status> TYPE ANY,
                 <lv_lock_owner> TYPE ANY,
                 <lv_lock_field> TYPE ANY,
                 <lv_lock_info> TYPE ANY,
                 <lv_row_message> TYPE ANY,
                 <lt_style> TYPE lvc_t_styl,
                 <lt_color> TYPE lvc_t_scol.
  DATA: ls_style TYPE lvc_s_styl,
        ls_lock  TYPE zflex_edit_lock,
        lv_locked_by_other TYPE abap_bool.

  CLEAR ls_lock.
  IF go_lock_mgr IS BOUND.
    ls_lock = go_lock_mgr->get_table_lock( ).
    lv_locked_by_other = go_lock_mgr->is_table_locked_by_other( ).
  ENDIF.

  LOOP AT pt_dyn_table ASSIGNING <dyn_wa>.
    UNASSIGN: <lt_style>, <lt_color>, <lv_is_new>, <lv_row_status>, <lv_lock_owner>, <lv_lock_field>, <lv_lock_info>, <lv_row_message>.
    ASSIGN COMPONENT 'CELL_STYLES' OF STRUCTURE <dyn_wa> TO <lt_style>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lt_style>.
    ENDIF.
    ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE <dyn_wa> TO <lt_color>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lt_color>.
    ENDIF.

    ASSIGN COMPONENT 'ROW_STATUS' OF STRUCTURE <dyn_wa> TO <lv_row_status>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_row_status>.
    ENDIF.

    ASSIGN COMPONENT 'LOCK_OWNER' OF STRUCTURE <dyn_wa> TO <lv_lock_owner>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_lock_owner>.
    ENDIF.
    ASSIGN COMPONENT 'LOCK_FIELD' OF STRUCTURE <dyn_wa> TO <lv_lock_field>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_lock_field>.
    ENDIF.
    ASSIGN COMPONENT 'LOCK_INFO' OF STRUCTURE <dyn_wa> TO <lv_lock_info>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_lock_info>.
    ENDIF.
    ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE <dyn_wa> TO <lv_row_message>. "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_row_message>.
    ENDIF.

    IF ls_lock-key_hash IS NOT INITIAL.
      IF <lv_lock_owner> IS ASSIGNED.
        <lv_lock_owner> = ls_lock-uname.
      ENDIF.
      IF <lv_lock_field> IS ASSIGNED.
        <lv_lock_field> = lcl_lock_manager=>gc_table_field.
      ENDIF.
      IF <lv_lock_info> IS ASSIGNED.
        IF ls_lock-uname = sy-uname.
          <lv_lock_info> = 'Table locked by me'. "#EC NOTEXT
        ELSE.
          <lv_lock_info> = |Table locked by { ls_lock-uname }|. "#EC NOTEXT
        ENDIF.
      ENDIF.
      IF lv_locked_by_other = abap_true.
        PERFORM set_row_status USING <dyn_wa> 'Failed'. "#EC NOTEXT
        DATA(lv_table_lock_msg) = |Table locked by { ls_lock-uname }|. "#EC NOTEXT
        PERFORM set_row_message USING <dyn_wa> lv_table_lock_msg.
      ENDIF.
    ENDIF.

    IF gv_edit_mode = abap_true AND <lt_style> IS ASSIGNED.
      ASSIGN COMPONENT 'IS_NEW_ROW' OF STRUCTURE <dyn_wa> TO <lv_is_new>. "#EC NOTEXT
      IF sy-subrc = 0 AND <lv_is_new> = abap_true.
        " New rows stay fully editable until the key is complete.
      ELSE.
        LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_fcat) WHERE key = 'X'. "#EC NOTEXT
          CLEAR ls_style.
          ls_style-fieldname = ls_fcat-fieldname.
          ls_style-style = cl_gui_alv_grid=>mc_style_disabled.
          INSERT ls_style INTO TABLE <lt_style>.
        ENDLOOP.
      ENDIF.

      IF lv_locked_by_other = abap_true.
        LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_lock_fcat)
          WHERE fieldname <> 'LOCK_OWNER' "#EC NOTEXT
            AND fieldname <> 'LOCK_FIELD' "#EC NOTEXT
            AND fieldname <> 'LOCK_INFO' "#EC NOTEXT
            AND fieldname <> 'ROW_MESSAGE'. "#EC NOTEXT
          CLEAR ls_style.
          ls_style-fieldname = ls_lock_fcat-fieldname.
          ls_style-style = cl_gui_alv_grid=>mc_style_disabled.
          INSERT ls_style INTO TABLE <lt_style>.
        ENDLOOP.
      ENDIF.
    ENDIF.
  ENDLOOP.
ENDFORM.

FORM load_gsheet_data.
  CONSTANTS:
    lc_default_dest TYPE rfcdest VALUE 'ZCPI_GOOGLE_SHEET', "#EC NOTEXT
    lc_default_path TYPE string  VALUE '/http/googlesheet/fetch'. "#EC NOTEXT

  DATA: lo_client TYPE REF TO if_http_client,
        lo_struct_desc TYPE REF TO cl_abap_structdescr,
        lo_table_desc  TYPE REF TO cl_abap_tabledescr,
        lt_comp        TYPE cl_abap_structdescr=>component_table,
        ls_comp        LIKE LINE OF lt_comp,
        lv_uri    TYPE string,
        lv_code   TYPE i,
        lv_reason TYPE string,
        lv_json   TYPE string.

  CLEAR: gv_gsheet_status,
         gv_gsheet_has_instance,
         gv_gsheet_can_export.
  IF gr_gsheet_table IS NOT BOUND.
    lo_struct_desc ?= cl_abap_typedescr=>describe_by_name( p_tabnam ).
    lt_comp = lo_struct_desc->get_components( ).

    READ TABLE lt_comp TRANSPORTING NO FIELDS WITH KEY name = 'CELL_COLORS'. "#EC NOTEXT
    IF sy-subrc <> 0.
      CLEAR ls_comp.
      ls_comp-name = 'CELL_COLORS'. "#EC NOTEXT
      ls_comp-type ?= cl_abap_typedescr=>describe_by_name( 'LVC_T_SCOL' ). "#EC NOTEXT
      APPEND ls_comp TO lt_comp.
    ENDIF.

    READ TABLE lt_comp TRANSPORTING NO FIELDS WITH KEY name = 'ROW_MESSAGE'. "#EC NOTEXT
    IF sy-subrc <> 0.
      CLEAR ls_comp.
      ls_comp-name = 'ROW_MESSAGE'. "#EC NOTEXT
      ls_comp-type = cl_abap_elemdescr=>get_c( 200 ).
      APPEND ls_comp TO lt_comp.
    ENDIF.

    lo_struct_desc = cl_abap_structdescr=>create( p_components = lt_comp ).
    lo_table_desc = cl_abap_tabledescr=>create( p_line_type = lo_struct_desc ).
    CREATE DATA gr_gsheet_table TYPE HANDLE lo_table_desc.
  ENDIF.

  UNASSIGN <gsheet_table>.
  ASSIGN gr_gsheet_table->* TO <gsheet_table>.
  IF sy-subrc <> 0 OR <gsheet_table> IS NOT ASSIGNED.
    gv_gsheet_status = |Google Sheet: cannot create dynamic table for { p_tabnam }|. "#EC NOTEXT
    RETURN.
  ENDIF.
  CLEAR <gsheet_table>.

  SELECT SINGLE spreadsheet_id, worksheet_name, range_name, cpi_dest, cpi_path
    FROM ztgsheet_map
    INTO @DATA(ls_map)
    WHERE tabname = @p_tabnam
      AND active = 'X'. "#EC NOTEXT

  IF sy-subrc <> 0.
    gv_gsheet_can_export = abap_true.
    gv_gsheet_status = |Google Sheet: table { p_tabnam } has no Cloud instance. Use Export to Cloud to create one.|. "#EC NOTEXT
    RETURN.
  ENDIF.

  IF ls_map-spreadsheet_id IS INITIAL OR ls_map-worksheet_name IS INITIAL.
    gv_gsheet_can_export = abap_true.
    gv_gsheet_status = |Google Sheet: Cloud instance mapping for { p_tabnam } is incomplete. Use Export to Cloud to create one.|. "#EC NOTEXT
    RETURN.
  ENDIF.

  gv_gsheet_has_instance = abap_true.

  DATA(lv_dest) = COND rfcdest( WHEN ls_map-cpi_dest IS INITIAL THEN lc_default_dest ELSE ls_map-cpi_dest ).
  lv_uri = COND string( WHEN ls_map-cpi_path IS INITIAL THEN lc_default_path ELSE ls_map-cpi_path ).

  DATA(lv_sheet_id) = CONV string( ls_map-spreadsheet_id ).
  DATA(lv_worksheet) = CONV string( ls_map-worksheet_name ).
  DATA(lv_range) = CONV string( ls_map-range_name ).

  PERFORM add_gsheet_query_param USING 'id' lv_sheet_id CHANGING lv_uri. "#EC NOTEXT
  PERFORM add_gsheet_query_param USING 'worksheetName' lv_worksheet CHANGING lv_uri. "#EC NOTEXT
  IF lv_range IS NOT INITIAL.
    PERFORM add_gsheet_query_param USING 'range' lv_range CHANGING lv_uri. "#EC NOTEXT
  ENDIF.

  cl_http_client=>create_by_destination(
    EXPORTING
      destination = lv_dest
    IMPORTING
      client      = lo_client
    EXCEPTIONS
      OTHERS      = 1 ).
  IF sy-subrc <> 0 OR lo_client IS NOT BOUND.
    gv_gsheet_status = |Google Sheet: cannot create HTTP destination { lv_dest }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  cl_http_utility=>set_request_uri(
    request = lo_client->request
    uri     = lv_uri ).

  lo_client->request->set_method( if_http_request=>co_request_method_get ).
  lo_client->request->set_header_field( name = 'Accept' value = 'application/json' ). "#EC NOTEXT

  lo_client->send( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    gv_gsheet_status = |Google Sheet: cannot send request to CPI|. "#EC NOTEXT
    lo_client->close( ).
    RETURN.
  ENDIF.

  lo_client->receive( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    gv_gsheet_status = |Google Sheet: cannot receive response from CPI|. "#EC NOTEXT
    lo_client->close( ).
    RETURN.
  ENDIF.

  lo_client->response->get_status(
    IMPORTING
      code   = lv_code
      reason = lv_reason ).

  lv_json = lo_client->response->get_cdata( ).
  lo_client->close( ).

  IF lv_code < 200 OR lv_code >= 300.
    gv_gsheet_status = |Google Sheet: CPI error HTTP { lv_code } { lv_reason }: { lv_json }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  TRY.
      /ui2/cl_json=>deserialize(
        EXPORTING
          json = lv_json
        CHANGING
          data = <gsheet_table> ).

      IF <gsheet_table> IS INITIAL.
        gv_gsheet_status = |Google Sheet: mapping found, but no rows returned for { p_tabnam }|. "#EC NOTEXT
      ELSE.
        gv_gsheet_status = |Google Sheet: { lines( <gsheet_table> ) } rows from worksheet { ls_map-worksheet_name }|. "#EC NOTEXT
      ENDIF.
      PERFORM compare_gsheet_with_ztab.
      PERFORM validate_gsheet_raw_json USING lv_json.
    CATCH cx_root INTO DATA(lx_gsheet_error).
      CLEAR <gsheet_table>.
      gv_gsheet_status = |Google Sheet: { lx_gsheet_error->get_text( ) }|. "#EC NOTEXT
  ENDTRY.
ENDFORM.

FORM add_gsheet_query_param USING pv_name  TYPE string
                                  pv_value TYPE string
                            CHANGING pv_uri TYPE string.
  DATA(lv_sep) = COND string( WHEN pv_uri CS '?' THEN '&' ELSE '?' ). "#EC NOTEXT
  DATA(lv_value) = cl_http_utility=>escape_url( pv_value ).

  pv_uri = |{ pv_uri }{ lv_sep }{ pv_name }={ lv_value }|. "#EC NOTEXT
ENDFORM.

FORM read_db_for_push CHANGING pt_push TYPE ANY TABLE.
  FIELD-SYMBOLS <lt_push> TYPE STANDARD TABLE.

  CLEAR pt_push.
  ASSIGN pt_push TO <lt_push>.
  IF sy-subrc <> 0 OR <lt_push> IS NOT ASSIGNED.
    gv_action_error = |Cannot assign push table for { p_tabnam }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  READ TABLE go_filter->mt_where_clauses INTO DATA(ls_where)
    WITH KEY tablename = p_tabnam.
  IF sy-subrc = 0 AND ls_where-where_tab IS NOT INITIAL.
    SELECT * FROM (p_tabnam) INTO CORRESPONDING FIELDS OF TABLE @<lt_push>
      WHERE (ls_where-where_tab).
  ELSE.
    SELECT * FROM (p_tabnam) INTO CORRESPONDING FIELDS OF TABLE @<lt_push>.
  ENDIF.
ENDFORM.

FORM build_sheet_payload USING pt_push TYPE ANY TABLE
                         CHANGING pv_payload   TYPE string
                                  pv_range     TYPE string
                                  pv_row_count TYPE i.
  lcl_gsheet_payload=>build_sheet_payload(
    EXPORTING it_push      = pt_push
    CHANGING  cv_payload   = pv_payload
              cv_range     = pv_range
              cv_row_count = pv_row_count ).
ENDFORM.

FORM call_cpi_create_sheet CHANGING pv_sheet_id  TYPE string
                                    pv_worksheet TYPE string
                                    pv_response  TYPE string.
  TYPES: BEGIN OF ty_create_sheet_prop,
           title TYPE string,
         END OF ty_create_sheet_prop.
  TYPES: BEGIN OF ty_create_sheet,
           properties TYPE ty_create_sheet_prop,
         END OF ty_create_sheet.
  TYPES tt_create_sheets TYPE STANDARD TABLE OF ty_create_sheet WITH EMPTY KEY.
  TYPES: BEGIN OF ty_create_response,
           spreadsheet_id TYPE string,
           sheets         TYPE tt_create_sheets,
         END OF ty_create_response.

  CONSTANTS:
    lc_default_dest TYPE rfcdest VALUE 'ZCPI_GOOGLE_SHEET', "#EC NOTEXT
    lc_create_path  TYPE string  VALUE '/http/googlesheet/create'. "#EC NOTEXT

  DATA: lo_client TYPE REF TO if_http_client,
        ls_create TYPE ty_create_response,
        lv_uri    TYPE string,
        lv_code   TYPE i,
        lv_reason TYPE string,
        lv_payload TYPE string,
        lv_title   TYPE string.

  CLEAR: pv_sheet_id,
         pv_worksheet,
         pv_response,
         gv_action_error.

  lv_uri = lc_create_path.
  lv_title = CONV string( p_tabnam ).
  PERFORM add_gsheet_query_param USING 'title' lv_title CHANGING lv_uri. "#EC NOTEXT
  lv_title = lcl_gsheet_payload=>json_escape( lv_title ).
  CONCATENATE '{"title":"' lv_title "#EC NOTEXT
              '","locale":"en_US","timeZone":"Asia/Bangkok","autoRecalc":"ON_CHANGE"}' "#EC NOTEXT
         INTO lv_payload.

  cl_http_client=>create_by_destination(
    EXPORTING
      destination = lc_default_dest
    IMPORTING
      client      = lo_client
    EXCEPTIONS
      OTHERS      = 1 ).
  IF sy-subrc <> 0 OR lo_client IS NOT BOUND.
    gv_action_error = |Google Sheet: cannot create HTTP destination { lc_default_dest }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  cl_http_utility=>set_request_uri(
    request = lo_client->request
    uri     = lv_uri ).

  lo_client->request->set_method( 'POST' ). "#EC NOTEXT
  lo_client->request->set_header_field( name = 'Accept' value = 'application/json' ). "#EC NOTEXT
  lo_client->request->set_header_field( name = 'Content-Type' value = 'application/json' ). "#EC NOTEXT
  lo_client->request->set_cdata( lv_payload ).

  lo_client->send( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    lo_client->close( ).
    gv_action_error = 'Google Sheet: cannot send create request to CPI'. "#EC NOTEXT
    RETURN.
  ENDIF.

  lo_client->receive( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    lo_client->close( ).
    gv_action_error = 'Google Sheet: cannot receive create response from CPI'. "#EC NOTEXT
    RETURN.
  ENDIF.

  lo_client->response->get_status(
    IMPORTING
      code   = lv_code
      reason = lv_reason ).

  pv_response = lo_client->response->get_cdata( ).
  lo_client->close( ).

  IF lv_code < 200 OR lv_code >= 300.
    gv_action_error = |CPI create error HTTP { lv_code } { lv_reason }: { pv_response }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  TRY.
      /ui2/cl_json=>deserialize(
        EXPORTING
          json        = pv_response
          pretty_name = /ui2/cl_json=>pretty_mode-camel_case
        CHANGING
          data        = ls_create ).
    CATCH cx_root INTO DATA(lx_create_json).
      gv_action_error = |Google Sheet: cannot parse create response: { lx_create_json->get_text( ) }|. "#EC NOTEXT
      RETURN.
  ENDTRY.

  pv_sheet_id = ls_create-spreadsheet_id.
  READ TABLE ls_create-sheets INTO DATA(ls_sheet) INDEX 1.
  IF sy-subrc = 0 AND ls_sheet-properties-title IS NOT INITIAL.
    pv_worksheet = ls_sheet-properties-title.
  ELSE.
    pv_worksheet = 'Sheet1'. "#EC NOTEXT
  ENDIF.

  IF pv_sheet_id IS INITIAL.
    gv_action_error = |Google Sheet: create response has no spreadsheetId: { pv_response }|. "#EC NOTEXT
    RETURN.
  ENDIF.
ENDFORM.

FORM upsert_gsheet_mapping USING pv_sheet_id  TYPE string
                                 pv_worksheet TYPE string.
  CONSTANTS:
    lc_default_dest TYPE rfcdest VALUE 'ZCPI_GOOGLE_SHEET', "#EC NOTEXT
    lc_fetch_path   TYPE string  VALUE '/http/googlesheet/fetch'. "#EC NOTEXT

  DATA: ls_map TYPE ztgsheet_map,
        lv_ts  TYPE timestampl.

  CLEAR gv_action_error.

  SELECT SINGLE client, tabname, spreadsheet_id, worksheet_name,
                range_name, cpi_dest, cpi_path, active,
                created_by, created_at, changed_by, changed_at
    FROM ztgsheet_map
    INTO CORRESPONDING FIELDS OF @ls_map
    WHERE tabname = @p_tabnam.

  IF sy-subrc <> 0.
    CLEAR ls_map.
    ls_map-tabname = p_tabnam.
    ls_map-created_by = sy-uname.
    ASSIGN COMPONENT 'CLIENT' OF STRUCTURE ls_map TO FIELD-SYMBOL(<lv_map_client>). "#EC NOTEXT
    IF sy-subrc <> 0.
      ASSIGN COMPONENT 'MANDT' OF STRUCTURE ls_map TO <lv_map_client>. "#EC NOTEXT
    ENDIF.
    IF sy-subrc = 0.
      <lv_map_client> = sy-mandt.
    ENDIF.
  ENDIF.

  GET TIME STAMP FIELD lv_ts.

  ls_map-spreadsheet_id = pv_sheet_id.
  ls_map-worksheet_name = pv_worksheet.
  CLEAR ls_map-range_name.
  ls_map-cpi_dest = lc_default_dest.
  ls_map-cpi_path = lc_fetch_path.
  ls_map-active = 'X'. "#EC NOTEXT
  ls_map-changed_by = sy-uname.

  ASSIGN COMPONENT 'CREATED_AT' OF STRUCTURE ls_map TO FIELD-SYMBOL(<lv_created_at>). "#EC NOTEXT
  IF sy-subrc = 0 AND <lv_created_at> IS INITIAL.
    <lv_created_at> = lv_ts.
  ENDIF.

  ASSIGN COMPONENT 'CHANGED_AT' OF STRUCTURE ls_map TO FIELD-SYMBOL(<lv_changed_at>). "#EC NOTEXT
  IF sy-subrc = 0.
    <lv_changed_at> = lv_ts.
  ENDIF.

  MODIFY ztgsheet_map FROM ls_map.
  IF sy-subrc <> 0.
    gv_action_error = |Google Sheet: cannot save mapping for { p_tabnam }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  COMMIT WORK AND WAIT.
ENDFORM.

FORM call_cpi_push USING pv_payload TYPE string
                         pv_range   TYPE string
                   CHANGING pv_response TYPE string.
  CONSTANTS:
    lc_default_dest TYPE rfcdest VALUE 'ZCPI_GOOGLE_SHEET', "#EC NOTEXT
    lc_push_path    TYPE string  VALUE '/http/googlesheet/update'. "#EC NOTEXT

  DATA: lo_client TYPE REF TO if_http_client,
        lv_uri    TYPE string,
        lv_code   TYPE i,
        lv_reason TYPE string.

  CLEAR pv_response.

  SELECT SINGLE spreadsheet_id, worksheet_name, range_name, cpi_dest
    FROM ztgsheet_map
    INTO @DATA(ls_map)
    WHERE tabname = @p_tabnam
      AND active = 'X'. "#EC NOTEXT

  IF sy-subrc <> 0.
    gv_action_error = |Google Sheet: table { p_tabnam } has no Cloud instance. Use Export to Cloud to create one.|. "#EC NOTEXT
    RETURN.
  ENDIF.

  IF ls_map-spreadsheet_id IS INITIAL OR ls_map-worksheet_name IS INITIAL.
    gv_action_error = |Google Sheet: mapping for { p_tabnam } is missing SheetID/WorkSheetName|. "#EC NOTEXT
    RETURN.
  ENDIF.

  DATA(lv_dest) = COND rfcdest( WHEN ls_map-cpi_dest IS INITIAL THEN lc_default_dest ELSE ls_map-cpi_dest ).
  DATA(lv_sheet_id) = CONV string( ls_map-spreadsheet_id ).
  DATA(lv_worksheet) = CONV string( ls_map-worksheet_name ).
  DATA(lv_range) = COND string( WHEN ls_map-range_name IS INITIAL THEN pv_range ELSE CONV string( ls_map-range_name ) ).

  lv_uri = lc_push_path.
  PERFORM add_gsheet_query_param USING 'id' lv_sheet_id CHANGING lv_uri. "#EC NOTEXT
  PERFORM add_gsheet_query_param USING 'worksheetName' lv_worksheet CHANGING lv_uri. "#EC NOTEXT
  PERFORM add_gsheet_query_param USING 'range' lv_range CHANGING lv_uri. "#EC NOTEXT

  cl_http_client=>create_by_destination(
    EXPORTING
      destination = lv_dest
    IMPORTING
      client      = lo_client
    EXCEPTIONS
      OTHERS      = 1 ).
  IF sy-subrc <> 0 OR lo_client IS NOT BOUND.
    gv_action_error = |Google Sheet: cannot create HTTP destination { lv_dest }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  cl_http_utility=>set_request_uri(
    request = lo_client->request
    uri     = lv_uri ).

  lo_client->request->set_method( 'PUT' ). "#EC NOTEXT
  lo_client->request->set_header_field( name = 'Accept' value = 'application/json' ). "#EC NOTEXT
  lo_client->request->set_header_field( name = 'Content-Type' value = 'application/json' ). "#EC NOTEXT
  lo_client->request->set_cdata( pv_payload ).

  lo_client->send( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    lo_client->close( ).
    gv_action_error = 'Google Sheet: cannot send push request to CPI'. "#EC NOTEXT
    RETURN.
  ENDIF.

  lo_client->receive( EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    lo_client->close( ).
    gv_action_error = 'Google Sheet: cannot receive push response from CPI'. "#EC NOTEXT
    RETURN.
  ENDIF.

  lo_client->response->get_status(
    IMPORTING
      code   = lv_code
      reason = lv_reason ).

  pv_response = lo_client->response->get_cdata( ).
  lo_client->close( ).

  IF lv_code < 200 OR lv_code >= 300.
    gv_action_error = |CPI push error HTTP { lv_code } { lv_reason }: { pv_response }|. "#EC NOTEXT
    RETURN.
  ENDIF.
ENDFORM.

FORM build_gsheet_fieldcat CHANGING pt_fieldcat TYPE lvc_t_fcat.
  CLEAR pt_fieldcat.

  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE' "#EC NOTEXT
    EXPORTING
      i_structure_name       = p_tabnam
    CHANGING
      ct_fieldcat            = pt_fieldcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.

  IF sy-subrc <> 0.
    gv_gsheet_status = |Google Sheet: cannot build field catalog for { p_tabnam }|. "#EC NOTEXT
    RETURN.
  ENDIF.

  LOOP AT pt_fieldcat ASSIGNING FIELD-SYMBOL(<ls_sheet_fcat>).
    <ls_sheet_fcat>-edit = space.
    <ls_sheet_fcat>-tech = space.
    <ls_sheet_fcat>-no_out = space.
    IF <ls_sheet_fcat>-fieldname = 'MANDT' "#EC NOTEXT
       OR <ls_sheet_fcat>-fieldname = 'CLIENT'. "#EC NOTEXT
      <ls_sheet_fcat>-tech = 'X'. "#EC NOTEXT
      <ls_sheet_fcat>-no_out = 'X'. "#EC NOTEXT
    ENDIF.
    IF <ls_sheet_fcat>-col_pos IS INITIAL.
      <ls_sheet_fcat>-col_pos = sy-tabix.
    ENDIF.
  ENDLOOP.

  DATA ls_msg_fcat TYPE lvc_s_fcat.
  CLEAR ls_msg_fcat.
  ls_msg_fcat-fieldname = 'ROW_MESSAGE'. "#EC NOTEXT
  ls_msg_fcat-coltext   = 'Compare Message'. "#EC NOTEXT
  ls_msg_fcat-scrtext_l = 'Compare Message'. "#EC NOTEXT
  ls_msg_fcat-scrtext_m = 'Message'. "#EC NOTEXT
  ls_msg_fcat-scrtext_s = 'Message'. "#EC NOTEXT
  ls_msg_fcat-outputlen = 60.
  ls_msg_fcat-edit      = space.
  APPEND ls_msg_fcat TO pt_fieldcat.
ENDFORM.

FORM validate_gsheet_raw_json USING pv_json TYPE string.
  DATA: lt_match    TYPE match_result_tab,
        ls_match    TYPE match_result,
        lv_object   TYPE string,
        lv_raw      TYPE string,
        lv_error    TYPE string,
        lv_message  TYPE string,
        lv_err_rows TYPE i.

  IF <gsheet_table> IS NOT ASSIGNED.
    RETURN.
  ENDIF.

  FIND ALL OCCURRENCES OF PCRE '\{[^{}]*\}' IN pv_json RESULTS lt_match.

  LOOP AT <gsheet_table> ASSIGNING FIELD-SYMBOL(<ls_sheet_row>).
    READ TABLE lt_match INTO ls_match INDEX sy-tabix.
    IF sy-subrc <> 0.
      EXIT.
    ENDIF.

    lv_object = pv_json+ls_match-offset(ls_match-length).

    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_raw_fcat).
      DATA(lv_is_tech) = abap_false.
      PERFORM is_technical_field USING ls_raw_fcat-fieldname CHANGING lv_is_tech.
      IF lv_is_tech = abap_true.
        CONTINUE.
      ENDIF.

      CLEAR: lv_raw, lv_error.
      PERFORM get_json_field_raw_value
        USING lv_object ls_raw_fcat-fieldname
        CHANGING lv_raw.

      IF lv_raw IS INITIAL.
        CONTINUE.
      ENDIF.

      PERFORM validate_cloud_raw_value
        USING ls_raw_fcat lv_raw
        CHANGING lv_error.

      IF lv_error IS NOT INITIAL.
        lv_err_rows = lv_err_rows + 1.
        lv_message = |Invalid Cloud value { ls_raw_fcat-fieldname }={ lv_raw }: { lv_error }|. "#EC NOTEXT
        PERFORM set_gsheet_message USING <ls_sheet_row> lv_message.
        PERFORM color_gsheet_cell USING <ls_sheet_row> ls_raw_fcat-fieldname 6.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

  IF lv_err_rows > 0.
    gv_gsheet_status = |{ gv_gsheet_status }; { lv_err_rows } invalid Cloud cell(s)|. "#EC NOTEXT
  ENDIF.
ENDFORM.

FORM get_json_field_raw_value USING pv_object    TYPE string
                                    pv_fieldname TYPE fieldname
                              CHANGING pv_value TYPE string.
  DATA: lv_pattern TYPE string,
        lv_field   TYPE string.

  CLEAR pv_value.
  lv_field = pv_fieldname.

  CONCATENATE '"' lv_field '"\s*:\s*"([^"]*)"' "#EC NOTEXT
    INTO lv_pattern.
  FIND PCRE lv_pattern IN pv_object SUBMATCHES pv_value.
  IF sy-subrc = 0.
    RETURN.
  ENDIF.

  CONCATENATE '"' lv_field '"\s*:\s*([^,}]*)' "#EC NOTEXT
    INTO lv_pattern.
  FIND PCRE lv_pattern IN pv_object SUBMATCHES pv_value.
  IF sy-subrc = 0.
    CONDENSE pv_value NO-GAPS.
  ENDIF.
ENDFORM.

FORM validate_cloud_raw_value USING ps_fcat TYPE lvc_s_fcat
                                    pv_raw  TYPE string
                              CHANGING pv_error TYPE string.
  DATA lv_value TYPE string.

  CLEAR pv_error.
  lv_value = pv_raw.
  CONDENSE lv_value NO-GAPS.

  IF lv_value IS INITIAL.
    RETURN.
  ENDIF.

  CASE ps_fcat-datatype.
    WHEN 'NUMC'. "#EC NOTEXT
      IF lv_value CN '0123456789'.
        pv_error = 'numeric text expected'. "#EC NOTEXT
      ENDIF.
    WHEN 'INT1' OR 'INT2' OR 'INT4' OR 'INT8'. "#EC NOTEXT
      IF lv_value CN '+-0123456789'.
        pv_error = 'integer expected'. "#EC NOTEXT
      ENDIF.
    WHEN 'DEC' OR 'CURR' OR 'QUAN' OR 'FLTP'. "#EC NOTEXT
      REPLACE ALL OCCURRENCES OF ',' IN lv_value WITH ''. "#EC NOTEXT
      REPLACE ALL OCCURRENCES OF '.' IN lv_value WITH ''. "#EC NOTEXT
      IF lv_value CN '+-0123456789'.
        pv_error = 'number expected'. "#EC NOTEXT
      ENDIF.
    WHEN 'DATS'. "#EC NOTEXT
      REPLACE ALL OCCURRENCES OF '-' IN lv_value WITH ''. "#EC NOTEXT
      IF strlen( lv_value ) <> 8 OR lv_value CN '0123456789'.
        pv_error = 'date YYYYMMDD expected'. "#EC NOTEXT
      ELSE.
        CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY' "#EC NOTEXT
          EXPORTING
            date                      = CONV sy-datum( lv_value )
          EXCEPTIONS
            plausibility_check_failed = 1
            OTHERS                    = 2.
        IF sy-subrc <> 0.
          pv_error = 'invalid date'. "#EC NOTEXT
        ENDIF.
      ENDIF.
    WHEN 'TIMS'. "#EC NOTEXT
      REPLACE ALL OCCURRENCES OF ':' IN lv_value WITH ''. "#EC NOTEXT
      IF strlen( lv_value ) <> 6 OR lv_value CN '0123456789'.
        pv_error = 'time HHMMSS expected'. "#EC NOTEXT
      ENDIF.
    WHEN OTHERS.
      IF ps_fcat-inttype = 'N' AND lv_value CN '0123456789'. "#EC NOTEXT
        pv_error = 'numeric text expected'. "#EC NOTEXT
      ENDIF.
  ENDCASE.
ENDFORM.

FORM normalize_cloud_value USING ps_fcat TYPE lvc_s_fcat
                                 pv_input TYPE string
                           CHANGING pv_value TYPE string.
  DATA: lv_len   TYPE i,
        lv_value TYPE string.

  lv_value = pv_input.

  CASE ps_fcat-datatype.
    WHEN 'DATS'. "#EC NOTEXT
      CONDENSE lv_value NO-GAPS.
      REPLACE ALL OCCURRENCES OF '-' IN lv_value WITH ''. "#EC NOTEXT
    WHEN 'TIMS'. "#EC NOTEXT
      CONDENSE lv_value NO-GAPS.
      REPLACE ALL OCCURRENCES OF ':' IN lv_value WITH ''. "#EC NOTEXT
    WHEN 'NUMC'. "#EC NOTEXT
      CONDENSE lv_value NO-GAPS.
      lv_len = ps_fcat-outputlen.
      IF lv_len IS INITIAL.
        lv_len = ps_fcat-intlen.
      ENDIF.
      WHILE lv_len > 0 AND strlen( lv_value ) < lv_len.
        lv_value = |0{ lv_value }|. "#EC NOTEXT
      ENDWHILE.
    WHEN OTHERS.
      IF ps_fcat-inttype = 'N'. "#EC NOTEXT
        CONDENSE lv_value NO-GAPS.
        lv_len = ps_fcat-outputlen.
        IF lv_len IS INITIAL.
          lv_len = ps_fcat-intlen.
        ENDIF.
        WHILE lv_len > 0 AND strlen( lv_value ) < lv_len.
          lv_value = |0{ lv_value }|. "#EC NOTEXT
        ENDWHILE.
      ENDIF.
  ENDCASE.

  pv_value = lv_value.
ENDFORM.

FORM is_technical_field USING pv_fieldname TYPE fieldname
                        CHANGING pv_is_tech TYPE abap_bool.
  pv_is_tech = lcl_field_util=>is_technical_field( pv_fieldname ).
ENDFORM.

FORM mark_cloud_preview_row USING ps_row     TYPE any
                                  pv_message TYPE string.
  lcl_alv_style=>mark_cloud_preview_row(
    EXPORTING iv_message = pv_message
    CHANGING  cs_row     = ps_row ).
ENDFORM.

FORM reload_gsheet_alv.
  IF go_grid IS BOUND.
    go_grid->check_changed_data( ).
  ENDIF.

  PERFORM load_gsheet_data.

  IF go_sheet_grid IS BOUND AND <gsheet_table> IS ASSIGNED.
    DATA lt_sheet_fcat TYPE lvc_t_fcat.
    PERFORM build_gsheet_fieldcat CHANGING lt_sheet_fcat.

    DATA(ls_sheet_layout) = VALUE lvc_s_layo(
      zebra      = 'X' "#EC NOTEXT
      sel_mode   = 'A' "#EC NOTEXT
      cwidth_opt = 'X' "#EC NOTEXT
      ctab_fname = 'CELL_COLORS' "#EC NOTEXT
      grid_title = gv_gsheet_status ).

    go_sheet_grid->set_frontend_fieldcatalog( it_fieldcatalog = lt_sheet_fcat ).
    go_sheet_grid->set_frontend_layout( is_layout = ls_sheet_layout ).
    go_sheet_grid->refresh_table_display( ).
    go_sheet_grid->set_toolbar_interactive( ).
  ENDIF.
ENDFORM.

FORM compare_gsheet_with_ztab.
  DATA: lt_key_fcat   TYPE lvc_t_fcat,
        lv_field_text TYPE string,
        lv_diff_count TYPE i,
        lv_missing_db TYPE i.

  FIELD-SYMBOLS <lt_db_missing_colors> TYPE lvc_t_scol.

  IF <gsheet_table> IS NOT ASSIGNED OR <dyn_table> IS NOT ASSIGNED.
    RETURN.
  ENDIF.

  LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_key_fcat).
    DATA(lv_skip_key_field) = abap_false.
    PERFORM is_cloud_compare_skip_field USING ls_key_fcat-fieldname CHANGING lv_skip_key_field.
    IF ls_key_fcat-key = 'X' "#EC NOTEXT
       AND lv_skip_key_field = abap_false.
      APPEND ls_key_fcat TO lt_key_fcat.
    ENDIF.
  ENDLOOP.

  LOOP AT <gsheet_table> ASSIGNING FIELD-SYMBOL(<ls_sheet_row>).
    ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE <ls_sheet_row> TO FIELD-SYMBOL(<lt_colors>). "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lt_colors>.
    ENDIF.

    ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE <ls_sheet_row> TO FIELD-SYMBOL(<lv_message>). "#EC NOTEXT
    IF sy-subrc = 0.
      CLEAR <lv_message>.
    ENDIF.

    DATA(lv_found) = abap_false.
    FIELD-SYMBOLS <ls_db_row> TYPE any.

    IF lt_key_fcat IS INITIAL.
      READ TABLE <dyn_table> INDEX sy-tabix ASSIGNING <ls_db_row>.
      IF sy-subrc = 0.
        lv_found = abap_true.
      ENDIF.
    ELSE.
      LOOP AT <dyn_table> ASSIGNING <ls_db_row>.
        DATA(lv_same_key) = abap_true.
        LOOP AT lt_key_fcat INTO DATA(ls_key_fcat_cmp).
          DATA(lv_sheet_key) = ||.
          DATA(lv_db_key) = ||.
          PERFORM get_compare_value USING <ls_sheet_row> ls_key_fcat_cmp CHANGING lv_sheet_key.
          PERFORM get_compare_value USING <ls_db_row>    ls_key_fcat_cmp CHANGING lv_db_key.
          IF lv_sheet_key <> lv_db_key.
            lv_same_key = abap_false.
            EXIT.
          ENDIF.
        ENDLOOP.
        IF lv_same_key = abap_true.
          lv_found = abap_true.
          EXIT.
        ENDIF.
      ENDLOOP.
    ENDIF.

    IF lv_found = abap_false.
      PERFORM color_gsheet_row USING <ls_sheet_row> 'New row in Google Sheet' 3. "#EC NOTEXT
      CONTINUE.
    ENDIF.

    CLEAR lv_diff_count.
    LOOP AT go_dyn_handler->mt_fieldcat INTO DATA(ls_cmp_fcat).
      DATA(lv_skip_cmp_field) = abap_false.
      PERFORM is_cloud_compare_skip_field USING ls_cmp_fcat-fieldname CHANGING lv_skip_cmp_field.
      IF lv_skip_cmp_field = abap_true.
        CONTINUE.
      ENDIF.

      DATA(lv_sheet_value) = ||.
      DATA(lv_db_value) = ||.
      PERFORM get_compare_value USING <ls_sheet_row> ls_cmp_fcat CHANGING lv_sheet_value.
      PERFORM get_compare_value USING <ls_db_row>    ls_cmp_fcat CHANGING lv_db_value.

      IF lv_sheet_value <> lv_db_value.
        lv_diff_count = lv_diff_count + 1.
        PERFORM color_gsheet_cell USING <ls_sheet_row> ls_cmp_fcat-fieldname 3.
      ENDIF.
    ENDLOOP.

    IF lv_diff_count > 0.
      lv_field_text = |Different from Z-table ({ lv_diff_count } field(s))|. "#EC NOTEXT
      PERFORM set_gsheet_message USING <ls_sheet_row> lv_field_text.
    ENDIF.
  ENDLOOP.

  LOOP AT <dyn_table> ASSIGNING FIELD-SYMBOL(<ls_db_missing>).
    DATA(lv_cloud_found) = abap_false.

    ASSIGN COMPONENT 'ROW_MESSAGE' OF STRUCTURE <ls_db_missing> TO FIELD-SYMBOL(<lv_db_missing_msg>). "#EC NOTEXT
    IF sy-subrc = 0 AND <lv_db_missing_msg> = 'Missing in Google Sheet'. "#EC NOTEXT
      CLEAR <lv_db_missing_msg>.
      ASSIGN COMPONENT 'CELL_COLORS' OF STRUCTURE <ls_db_missing> TO <lt_db_missing_colors>. "#EC NOTEXT
      IF sy-subrc = 0.
        DELETE <lt_db_missing_colors> WHERE color-col = 3.
      ENDIF.
    ENDIF.

    IF lt_key_fcat IS INITIAL.
      READ TABLE <gsheet_table> INDEX sy-tabix TRANSPORTING NO FIELDS.
      IF sy-subrc = 0.
        lv_cloud_found = abap_true.
      ENDIF.
    ELSE.
      LOOP AT <gsheet_table> ASSIGNING FIELD-SYMBOL(<ls_sheet_cmp>).
        DATA(lv_missing_same_key) = abap_true.
        LOOP AT lt_key_fcat INTO DATA(ls_missing_key_fcat).
          DATA(lv_missing_db_key) = ||.
          DATA(lv_missing_sheet_key) = ||.
          PERFORM get_compare_value USING <ls_db_missing>  ls_missing_key_fcat CHANGING lv_missing_db_key.
          PERFORM get_compare_value USING <ls_sheet_cmp>   ls_missing_key_fcat CHANGING lv_missing_sheet_key.
          IF lv_missing_db_key <> lv_missing_sheet_key.
            lv_missing_same_key = abap_false.
            EXIT.
          ENDIF.
        ENDLOOP.
        IF lv_missing_same_key = abap_true.
          lv_cloud_found = abap_true.
          EXIT.
        ENDIF.
      ENDLOOP.
    ENDIF.

    IF lv_cloud_found = abap_false.
      lv_missing_db = lv_missing_db + 1.
      PERFORM color_gsheet_row USING <ls_db_missing> 'Missing in Google Sheet' 3. "#EC NOTEXT
    ENDIF.
  ENDLOOP.

  IF lv_missing_db > 0.
    gv_gsheet_status = |{ gv_gsheet_status }; { lv_missing_db } Z-table row(s) missing in Cloud|. "#EC NOTEXT
  ENDIF.

  IF go_grid IS BOUND.
    go_grid->refresh_table_display( ).
  ENDIF.
ENDFORM.

FORM build_cloud_compare_key USING ps_row      TYPE any
                                   pt_key_fcat TYPE lvc_t_fcat
                                   pv_index    TYPE sy-tabix
                             CHANGING pv_key TYPE string.
  pv_key = lcl_field_util=>build_cloud_compare_key(
    is_row      = ps_row
    it_key_fcat = pt_key_fcat
    iv_index    = pv_index ).
ENDFORM.

FORM fill_generated_technical_keys USING ps_row TYPE any.
  lcl_field_util=>fill_generated_technical_keys(
    CHANGING cs_row = ps_row ).
ENDFORM.

FORM is_cloud_compare_skip_field USING pv_fieldname TYPE fieldname
                                 CHANGING pv_skip TYPE abap_bool.
  pv_skip = lcl_field_util=>is_cloud_compare_skip_field( pv_fieldname ).
ENDFORM.

FORM get_compare_value USING ps_row  TYPE any
                             ps_fcat TYPE lvc_s_fcat
                       CHANGING pv_value TYPE string.
  pv_value = lcl_field_util=>get_compare_value(
    is_row  = ps_row
    is_fcat = ps_fcat ).
ENDFORM.

FORM color_gsheet_row USING ps_row     TYPE any
                            pv_message TYPE string
                            pv_color   TYPE i.
  lcl_alv_style=>color_gsheet_row(
    EXPORTING
      iv_message = pv_message
      iv_color   = pv_color
    CHANGING
      cs_row     = ps_row ).
ENDFORM.

FORM color_gsheet_cell USING ps_row       TYPE any
                             pv_fieldname TYPE fieldname
                             pv_color     TYPE i.
  lcl_alv_style=>color_gsheet_cell(
    EXPORTING
      iv_fieldname = pv_fieldname
      iv_color     = pv_color
    CHANGING
      cs_row       = ps_row ).
ENDFORM.

FORM recolor_row_cells USING ps_row        TYPE any
                             pv_from_color TYPE i
                             pv_to_color   TYPE i.
  lcl_alv_style=>recolor_row_cells(
    EXPORTING
      iv_from_color = pv_from_color
      iv_to_color   = pv_to_color
    CHANGING
      cs_row        = ps_row ).
ENDFORM.

FORM set_gsheet_message USING ps_row     TYPE any
                              pv_message TYPE string.
  lcl_alv_style=>set_gsheet_message(
    EXPORTING iv_message = pv_message
    CHANGING  cs_row     = ps_row ).
ENDFORM.

START-OF-SELECTION.
  IF p_maint = 'X'. "#EC NOTEXT
    gv_edit_mode = abap_true.
  ELSE.
    gv_edit_mode = abap_false.
  ENDIF.

  CREATE OBJECT go_dyn_handler.
  CREATE OBJECT go_filter.
  CREATE OBJECT go_excel_sync.
  CREATE OBJECT go_lock_mgr.
  CREATE OBJECT go_alv_report.
  CREATE OBJECT go_audit.

  go_dyn_handler->create_dynamic_table( ).
  go_dyn_handler->build_fieldcatalog( ).

  " Initialize dynamic filter (SE16N style)
  go_filter->initialize( ).

  " Execute SELECT with or without filter
  go_filter->execute_select( ).

  " Reset user commands before displaying host screen 2000 to prevent WebGUI from immediately exiting.
  CLEAR: sy-ucomm, sscrfields-ucomm.
  CALL SELECTION-SCREEN 2000.
