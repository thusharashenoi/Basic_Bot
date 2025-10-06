************************************************************************
*  Confidential property of PepsiCo                                     *
*  All Rights Reserved                                                  *
************************************************************************
*  Report Name   : ZLREP_CONDITION_CONTRACT_TRS                         *
*  Created on    : 26/09/2025                                           *
*  RICEF         :                                                      *
*  Project       : MENA+                                                *
*  Description   : End-to-end create/change of Condition Contracts in   *
*                  bulk with upload/download templates and ALV results  *
*                  - Robust commit/rollback handling                    *
*                  - Correct X-flag usage for BAPI changes              *
*                  - Currency handling (CURRENCY / CURRENCY_ISO)        *
*                  - No-change detection to avoid false successes       *
*                  - Excel parsing with date normalization              *
*  Version       : 2.0                                                  *
************************************************************************

REPORT zlrep_condition_contract_trs.

* Text Symbols
* text-001: Condition Contract Action
* text-002: Template Handling
* text-003: File Path for Upload/Download
* text-004: Method of Execution
* text-005: Test Run

* Selection Screen Definitions

SELECTION-SCREEN BEGIN OF BLOCK blk1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_create RADIOBUTTON GROUP grp1 DEFAULT 'X',
              p_change RADIOBUTTON GROUP grp1.
SELECTION-SCREEN END OF BLOCK blk1.

SELECTION-SCREEN BEGIN OF BLOCK blk2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_upload RADIOBUTTON GROUP grp2 DEFAULT 'X' USER-COMMAND cmd,
              p_down   RADIOBUTTON GROUP grp2.
SELECTION-SCREEN END OF BLOCK blk2.

SELECTION-SCREEN BEGIN OF BLOCK blk3 WITH FRAME TITLE TEXT-003.
  PARAMETERS: p_file TYPE string OBLIGATORY.
SELECTION-SCREEN END OF BLOCK blk3.

SELECTION-SCREEN BEGIN OF BLOCK blk4 WITH FRAME TITLE TEXT-004.
  PARAMETERS: p_test AS CHECKBOX.
SELECTION-SCREEN END OF BLOCK blk4.


CLASS lcl_condition_contract DEFINITION FINAL CREATE PRIVATE.
  PUBLIC SECTION.
    TYPES : BEGIN OF lty_condition_data,
              contract_number    TYPE char10, "For change
              currency           TYPE waers, "Change: Contract Currency
              contract_type      TYPE char4,
              process_variant    TYPE char4,
              ext_num            TYPE char30,
              cust_owner         TYPE char10,
              date_from          TYPE dats,
              date_to            TYPE dats,
              vkorg              TYPE char4,
              vtweg              TYPE char2,
              spart              TYPE char2,
              zterm              TYPE char4,
              settl_cal_accr     TYPE char2,
              assignment         TYPE char16,
              settl_cal_part     TYPE char2,
              settl_cal_delta    TYPE char2,
              cc_curr            TYPE waers,
              exchange_rate      TYPE char9,
              exchange_rate_type TYPE char4,
              " Business Volume Block fields
              order_key          TYPE char10,
              fieldcomb          TYPE char4,
              incl_excl          TYPE char1,
              kunnr              TYPE char10,
              matnr              TYPE char40,
              vkorg_bvb          TYPE char4,
              augru              TYPE char3,
              kunhier            TYPE char10,
              prodh              TYPE char18,
              mvgr1              TYPE char3,
              mvgr2              TYPE char3,
              mvgr3              TYPE char3,
              mvgr4              TYPE char3,
              mvgr5              TYPE char3,
              kvgr1              TYPE char3,
              kvgr2              TYPE char3,
              kvgr3              TYPE char3,
              kvgr4              TYPE char3,
              kvgr5              TYPE char3,
              prodh1             TYPE char2,
              prodh2             TYPE char2,
              prodh3             TYPE char3,
              zzprodh4           TYPE char3,
              zzprodh5           TYPE char3,
              zzprodh6           TYPE char2,
              zzprodh7           TYPE char3,
              vtweg_bvb          TYPE char2,
              werks              TYPE char4,
              spart_bvb          TYPE char2,
              kunrg              TYPE char10,
            END OF lty_condition_data,
            BEGIN OF lty_message,
              row_number         TYPE i,
              contract_type      TYPE char4,
              process_variant    TYPE char4,
              ext_num            TYPE char30,
              cust_owner         TYPE char10,
              date_from          TYPE dats,
              date_to            TYPE dats,
              contract_number    TYPE char10,
              order_key          TYPE char10,
              currency           TYPE waers,
              type               TYPE symsgty,
              message            TYPE string,
              validation_step    TYPE string,
            END OF lty_message,
            BEGIN OF lty_validation_result,
              valid              TYPE abap_bool,
              error_message      TYPE string,
              validation_step    TYPE string,
            END OF lty_validation_result.

    CLASS-METHODS get_instance
      RETURNING
        VALUE(ro_instance) TYPE REF TO lcl_condition_contract.

    METHODS:
      authorization_check,
      download_template,
      check_input_file,
      check_file_exist,
      get_data,
      create_condition_contracts,
      change_condition_contracts,
      display_results,
      selection_screen,
      has_data RETURNING VALUE(rv_has_data) TYPE abap_bool.

  PRIVATE SECTION.
    DATA: gt_condition_data TYPE  TABLE OF lty_condition_data WITH EMPTY KEY,
          gt_message        TYPE STANDARD TABLE OF lty_message,
          gv_success_count  TYPE i,
          gv_error_count    TYPE i.

    CLASS-DATA:
      go_instance TYPE REF TO lcl_condition_contract.

    CONSTANTS: gc_success TYPE string VALUE 'S',
               gc_error   TYPE string VALUE 'E',
               gc_warning TYPE string VALUE 'W',
               gc_abort   TYPE string VALUE 'A'.

    METHODS:
      validate_data_create
        IMPORTING
          is_data          TYPE lty_condition_data
          iv_row_number    TYPE i
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      validate_data_change
        IMPORTING
          is_data          TYPE lty_condition_data
          iv_row_number    TYPE i
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      validate_mandatory_fields
        IMPORTING
          is_data          TYPE lty_condition_data
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      validate_dates
        IMPORTING
          is_data          TYPE lty_condition_data
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      format_customer
        IMPORTING
          iv_customer      TYPE char10
        RETURNING
          VALUE(rv_kunnr)  TYPE kna1-kunnr,

      call_bapi_create
        IMPORTING
          is_data          TYPE lty_condition_data
          iv_row_number    TYPE i
        RETURNING
          VALUE(rs_result) TYPE lty_message,

      call_bapi_change
        IMPORTING
          is_data          TYPE lty_condition_data
          iv_row_number    TYPE i
        RETURNING
          VALUE(rs_result) TYPE lty_message,

      set_alv_column_headers
        IMPORTING io_alv TYPE REF TO cl_salv_table,

      convert_excel_date
        IMPORTING
          iv_excel_date      TYPE any
        RETURNING
          VALUE(rv_sap_date) TYPE dats,

      validate_customer
        IMPORTING
          is_data          TYPE lty_condition_data
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      validate_data_enhanced
        IMPORTING
          is_data          TYPE lty_condition_data
          iv_row_number    TYPE i
        RETURNING
          VALUE(rs_result) TYPE lty_validation_result,

      call_bapi_create_group
        IMPORTING
          is_header          TYPE lty_condition_data
          it_bvb_rows        TYPE STANDARD TABLE
          iv_header_row      TYPE i
        RETURNING
          VALUE(rs_result)   TYPE lty_message.

ENDCLASS.


CLASS lcl_condition_contract IMPLEMENTATION.

  METHOD get_instance.
    IF go_instance IS NOT BOUND.
      go_instance = NEW #( ).
    ENDIF.
    ro_instance = go_instance.
  ENDMETHOD.

  METHOD authorization_check.
    AUTHORITY-CHECK OBJECT 'S_TCODE'
    ID 'TCD' FIELD sy-tcode.
    IF sy-subrc NE 0.
      MESSAGE 'No authorization for this transaction' TYPE 'E'.
    ENDIF.
  ENDMETHOD.

  METHOD convert_excel_date.
    DATA: lv_input_string TYPE string,
          lv_year         TYPE string,
          lv_month        TYPE string,
          lv_day          TYPE string,
          lv_num          TYPE i.

    lv_input_string = |{ iv_excel_date }|.
    CONDENSE lv_input_string NO-GAPS.

    IF lv_input_string IS INITIAL.
      RETURN.
    ENDIF.

    " Excel serial number handling (base 1899-12-30)
    IF lv_input_string CO '0123456789'.
      lv_num = lv_input_string.
      IF lv_num > 0 AND lv_num < 80000 AND strlen( lv_input_string ) < 8.
        rv_sap_date = '18991230' + lv_num.
      ELSEIF strlen( lv_input_string ) = 8.
        rv_sap_date = lv_input_string.
      ENDIF.
    ELSEIF lv_input_string CA '/'.
      SPLIT lv_input_string AT '/' INTO TABLE DATA(lt_date_parts).
      IF lines( lt_date_parts ) = 3.
        lv_month = |{ lt_date_parts[ 1 ] ALPHA = IN }|.
        lv_day   = |{ lt_date_parts[ 2 ] ALPHA = IN }|.
        lv_year  = lt_date_parts[ 3 ].
        rv_sap_date = |{ lv_year }{ lv_month }{ lv_day }|.
      ENDIF.
    ELSEIF lv_input_string CA '-'.
      SPLIT lv_input_string AT '-' INTO TABLE lt_date_parts.
      IF lines( lt_date_parts ) = 3.
        lv_year  = lt_date_parts[ 1 ].
        lv_month = |{ lt_date_parts[ 2 ] ALPHA = IN }|.
        lv_day   = |{ lt_date_parts[ 3 ] ALPHA = IN }|.
        rv_sap_date = |{ lv_year }{ lv_month }{ lv_day }|.
      ENDIF.
    ENDIF.

    IF rv_sap_date IS INITIAL.
      rv_sap_date = lv_input_string.
    ENDIF.

    CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
      EXPORTING
        date                      = rv_sap_date
      EXCEPTIONS
        plausibility_check_failed = 1
        OTHERS                    = 2.
    IF sy-subrc <> 0.
      CLEAR rv_sap_date.
    ENDIF.
  ENDMETHOD.

  METHOD validate_mandatory_fields.
    rs_result-valid = abap_true.
    rs_result-validation_step = 'MANDATORY_FIELDS'.

    IF is_data-contract_type IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Contract Type is mandatory and cannot be empty'.
      RETURN.
    ENDIF.

    IF is_data-process_variant IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Process Variant is mandatory and cannot be empty'.
      RETURN.
    ENDIF.

    IF is_data-cust_owner IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Customer Owner is mandatory and cannot be empty'.
      RETURN.
    ENDIF.

    IF is_data-date_from IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Valid From Date is mandatory and cannot be empty'.
      RETURN.
    ENDIF.

    IF is_data-date_to IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Valid To Date is mandatory and cannot be empty'.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD validate_dates.
    rs_result-valid = abap_true.
    rs_result-validation_step = 'DATE_VALIDATION'.

    IF is_data-date_from IS INITIAL OR is_data-date_to IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Invalid date format detected'.
      RETURN.
    ENDIF.

    IF is_data-date_from > is_data-date_to.
      rs_result-valid = abap_false.
      rs_result-error_message = |Valid From ({ is_data-date_from }) cannot be greater than Valid To ({ is_data-date_to })|.
      RETURN.
    ENDIF.

    IF is_data-date_to < sy-datum.
      rs_result-valid = abap_false.
      rs_result-error_message = |Valid To ({ is_data-date_to }) cannot be in the past|.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD format_customer.
    DATA lv_customer TYPE kna1-kunnr.
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING input  = iv_customer
      IMPORTING output = lv_customer.
    rv_kunnr = lv_customer.
  ENDMETHOD.

  METHOD validate_customer.
    rs_result-valid = abap_true.
    rs_result-validation_step = 'CUSTOMER_VALIDATION'.

    DATA: lv_customer TYPE kna1-kunnr.

    IF is_data-cust_owner IS NOT INITIAL.
      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
        EXPORTING input  = is_data-cust_owner
        IMPORTING output = lv_customer.

      SELECT SINGLE kunnr FROM kna1 INTO lv_customer WHERE kunnr = lv_customer.
      IF sy-subrc <> 0.
        rs_result-valid = abap_false.
        rs_result-error_message = |Customer { is_data-cust_owner } does not exist in system|.
        RETURN.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD validate_data_enhanced.
    DATA ls_result TYPE lty_validation_result.

    ls_result = validate_mandatory_fields( is_data ).
    IF ls_result-valid = abap_false.
      rs_result = ls_result.
      RETURN.
    ENDIF.

    ls_result = validate_dates( is_data ).
    IF ls_result-valid = abap_false.
      rs_result = ls_result.
      RETURN.
    ENDIF.

    ls_result = validate_customer( is_data ).
    IF ls_result-valid = abap_false.
      rs_result = ls_result.
      RETURN.
    ENDIF.

    rs_result-valid = abap_true.
    rs_result-validation_step = 'ALL_VALIDATIONS_PASSED'.
  ENDMETHOD.

  METHOD validate_data_change.
    rs_result-valid = abap_true.
    rs_result-validation_step = 'CHANGE_VALIDATION'.

    IF is_data-contract_number IS INITIAL.
      rs_result-valid = abap_false.
      rs_result-error_message = 'Condition Contract Number is mandatory for change'.
      RETURN.
    ENDIF.

    " Dates are optional for change, but if supplied validate order
    IF is_data-date_from IS NOT INITIAL AND is_data-date_to IS NOT INITIAL.
      IF is_data-date_from > is_data-date_to.
        rs_result-valid = abap_false.
        rs_result-error_message = |Valid From ({ is_data-date_from }) cannot be greater than Valid To ({ is_data-date_to })|.
        RETURN.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD call_bapi_create.
    DATA lt_bvb TYPE STANDARD TABLE OF lty_condition_data WITH EMPTY KEY.
    APPEND is_data TO lt_bvb.
    rs_result = call_bapi_create_group( is_header = is_data it_bvb_rows = lt_bvb iv_header_row = iv_row_number ).
  ENDMETHOD.

  METHOD call_bapi_create_group.
    DATA: ls_headdatain         TYPE bapicchead,
          ls_headdatainx        TYPE bapiccheadx,
          lv_contract_number    TYPE bapicckey-condition_contract_number,
          lt_return             TYPE TABLE OF bapiret2,
          ls_return             TYPE bapiret2,
          ls_commit_return      TYPE bapiret2,
          lv_customer_formatted TYPE kna1-kunnr,
          lv_all_messages       TYPE string,
          lv_separator          TYPE string VALUE '; '.

    CLEAR rs_result.
    rs_result-row_number      = iv_header_row.
    rs_result-contract_type   = is_header-contract_type.
    rs_result-process_variant = is_header-process_variant.
    rs_result-ext_num         = is_header-ext_num.
    rs_result-cust_owner      = is_header-cust_owner.
    rs_result-date_from       = is_header-date_from.
    rs_result-date_to         = is_header-date_to.

    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING input  = is_header-cust_owner
      IMPORTING output = lv_customer_formatted.

    DATA: lt_bapiccbvb  TYPE TABLE OF bapiccbvb,
          lt_bapiccbvbx TYPE TABLE OF bapiccbvbx,
          ls_bapiccbvb  TYPE bapiccbvb,
          ls_bapiccbvbx TYPE bapiccbvbx.

    DATA(lv_idx) = 0.
    LOOP AT it_bvb_rows ASSIGNING FIELD-SYMBOL(<ls_bvb_row>).
      CLEAR: ls_bapiccbvb, ls_bapiccbvbx.
      ADD 1 TO lv_idx.
      DATA(lv_key) = COND char10( WHEN <ls_bvb_row>-order_key IS NOT INITIAL THEN <ls_bvb_row>-order_key ELSE |{ lv_idx }| ).

      ls_bapiccbvb-order_key        = lv_key.
      ls_bapiccbvb-include_exclude  = COND #( WHEN <ls_bvb_row>-incl_excl IS NOT INITIAL THEN <ls_bvb_row>-incl_excl ELSE 'I' ).
      IF <ls_bvb_row>-fieldcomb IS NOT INITIAL.
        ls_bapiccbvb-fieldcombination = <ls_bvb_row>-fieldcomb.
      ENDIF.

      IF <ls_bvb_row>-kunnr IS NOT INITIAL.
        CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
          EXPORTING input  = <ls_bvb_row>-kunnr
          IMPORTING output = ls_bapiccbvb-customer_new.
        ls_bapiccbvb-customer_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-matnr IS NOT INITIAL.
        ls_bapiccbvb-material_new = <ls_bvb_row>-matnr.
        ls_bapiccbvb-material_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-vkorg_bvb IS NOT INITIAL.
        ls_bapiccbvb-salesorg_new = <ls_bvb_row>-vkorg_bvb.
        ls_bapiccbvb-salesorg_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-augru IS NOT INITIAL.
        ls_bapiccbvb-ord_reason_new = <ls_bvb_row>-augru.
        ls_bapiccbvb-ord_reason_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-kunhier IS NOT INITIAL.
        ls_bapiccbvb-cust_hier_new = <ls_bvb_row>-kunhier.
        ls_bapiccbvb-cust_hier_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-prodh IS NOT INITIAL.
        ls_bapiccbvb-prod_hier_new = <ls_bvb_row>-prodh.
        ls_bapiccbvb-prod_hier_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-kvgr1 IS NOT INITIAL.
        ls_bapiccbvb-cust_grp1_new = <ls_bvb_row>-kvgr1.
        ls_bapiccbvb-cust_grp1_key = lv_key.
      ENDIF.
      IF <ls_bvb_row>-kvgr2 IS NOT INITIAL.
        ls_bapiccbvb-cust_grp2_new = <ls_bvb_row>-kvgr2.
        ls_bapiccbvb-cust_grp2_key = lv_key.
      ENDIF.
      IF <ls_bvb_row>-kvgr3 IS NOT INITIAL.
        ls_bapiccbvb-cust_grp3_new = <ls_bvb_row>-kvgr3.
        ls_bapiccbvb-cust_grp3_key = lv_key.
      ENDIF.
      IF <ls_bvb_row>-kvgr4 IS NOT INITIAL.
        ls_bapiccbvb-cust_grp4_new = <ls_bvb_row>-kvgr4.
        ls_bapiccbvb-cust_grp4_key = lv_key.
      ENDIF.
      IF <ls_bvb_row>-kvgr5 IS NOT INITIAL.
        ls_bapiccbvb-cust_grp5_new = <ls_bvb_row>-kvgr5.
        ls_bapiccbvb-cust_grp5_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-vtweg_bvb IS NOT INITIAL.
        ls_bapiccbvb-distr_chan_new = <ls_bvb_row>-vtweg_bvb.
        ls_bapiccbvb-distr_chan_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-werks IS NOT INITIAL.
        ls_bapiccbvb-plant_new = <ls_bvb_row>-werks.
        ls_bapiccbvb-plant_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-spart_bvb IS NOT INITIAL.
        ls_bapiccbvb-division_new = <ls_bvb_row>-spart_bvb.
        ls_bapiccbvb-division_key = lv_key.
      ENDIF.

      IF <ls_bvb_row>-kunrg IS NOT INITIAL.
        CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
          EXPORTING input  = <ls_bvb_row>-kunrg
          IMPORTING output = ls_bapiccbvb-payer_new.
        ls_bapiccbvb-payer_key = lv_key.
      ENDIF.

      APPEND ls_bapiccbvb TO lt_bapiccbvb.

      ls_bapiccbvbx-order_key  = lv_key.
      ls_bapiccbvbx-updateflag = 'U'.
      APPEND ls_bapiccbvbx TO lt_bapiccbvbx.
    ENDLOOP.

    ls_headdatain-contract_type       = is_header-contract_type.
    ls_headdatain-process_variant     = is_header-process_variant.
    ls_headdatain-customer_owner      = lv_customer_formatted.
    ls_headdatain-validity_date_from  = is_header-date_from.
    ls_headdatain-validity_date_to    = is_header-date_to.
    ls_headdatain-salesorg            = is_header-vkorg.
    ls_headdatain-distr_chan          = is_header-vtweg.
    ls_headdatain-division            = is_header-spart.

    IF is_header-exchange_rate IS NOT INITIAL.
      ls_headdatain-exch_rate = is_header-exchange_rate.
    ENDIF.
    IF is_header-exchange_rate_type IS NOT INITIAL.
      ls_headdatain-exchg_rate = is_header-exchange_rate_type.
    ENDIF.

    ls_headdatainx-contract_type       = 'X'.
    ls_headdatainx-process_variant     = 'X'.
    ls_headdatainx-customer_owner      = 'X'.
    ls_headdatainx-validity_date_from  = 'X'.
    ls_headdatainx-validity_date_to    = 'X'.
    ls_headdatainx-salesorg            = 'X'.
    ls_headdatainx-distr_chan          = 'X'.
    ls_headdatainx-division            = 'X'.

    IF is_header-zterm IS NOT INITIAL.
      ls_headdatainx-pmnttrms = 'X'.
    ENDIF.
    IF is_header-assignment IS NOT INITIAL.
      ls_headdatainx-assignment = 'X'.
    ENDIF.
    IF is_header-exchange_rate IS NOT INITIAL.
      ls_headdatainx-exch_rate = 'X'.
    ENDIF.
    IF is_header-exchange_rate_type IS NOT INITIAL.
      ls_headdatainx-exchg_rate = 'X'.
    ENDIF.

    CALL FUNCTION 'BAPI_CONDITION_CONTRACT_CREATE'
      EXPORTING
        headdatain              = ls_headdatain
        headdatainx             = ls_headdatainx
      IMPORTING
        conditioncontractnumber = lv_contract_number
      TABLES
        bvbdatain               = lt_bapiccbvb
        bvbdatainx              = lt_bapiccbvbx
        return                  = lt_return.

    DATA(lv_detail) = VALUE string( ).
    CLEAR lv_all_messages.
    LOOP AT lt_return INTO ls_return WHERE type = gc_abort.
      lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
      IF ls_return-message_v1 IS NOT INITIAL OR
         ls_return-message_v2 IS NOT INITIAL OR
         ls_return-message_v3 IS NOT INITIAL OR
         ls_return-message_v4 IS NOT INITIAL.
        lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
      ENDIF.
      IF lv_all_messages IS INITIAL.
        lv_all_messages = |Abort: { lv_detail }|.
      ELSE.
        lv_all_messages = |{ lv_all_messages }{ lv_separator }Abort: { lv_detail }|.
      ENDIF.
    ENDLOOP.

    IF lv_all_messages IS NOT INITIAL.
      rs_result-type = gc_error.
      rs_result-validation_step = 'BAPI_ABORT'.
      rs_result-message = lv_all_messages.
      RETURN.
    ENDIF.

    CLEAR lv_all_messages.
    LOOP AT lt_return INTO ls_return WHERE type = gc_error.
      lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
      IF ls_return-message_v1 IS NOT INITIAL OR
         ls_return-message_v2 IS NOT INITIAL OR
         ls_return-message_v3 IS NOT INITIAL OR
         ls_return-message_v4 IS NOT INITIAL.
        lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
      ENDIF.
      IF lv_all_messages IS INITIAL.
        lv_all_messages = |Error: { lv_detail }|.
      ELSE.
        lv_all_messages = |{ lv_all_messages }{ lv_separator }Error: { lv_detail }|.
      ENDIF.
    ENDLOOP.

    IF lv_all_messages IS NOT INITIAL.
      rs_result-type = gc_error.
      rs_result-validation_step = 'BAPI_ERROR'.
      rs_result-message = lv_all_messages.
      RETURN.
    ENDIF.

    READ TABLE lt_return INTO ls_return WITH KEY type = gc_success.
    IF sy-subrc = 0 OR lv_contract_number IS NOT INITIAL.
      CLEAR lv_all_messages.
      LOOP AT lt_return INTO ls_return WHERE type = gc_warning.
        lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
        IF ls_return-message_v1 IS NOT INITIAL OR
           ls_return-message_v2 IS NOT INITIAL OR
           ls_return-message_v3 IS NOT INITIAL OR
           ls_return-message_v4 IS NOT INITIAL.
          lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
        ENDIF.
        IF lv_all_messages IS INITIAL.
          lv_all_messages = |Warning: { lv_detail }|.
        ELSE.
          lv_all_messages = |{ lv_all_messages }{ lv_separator }Warning: { lv_detail }|.
        ENDIF.
      ENDLOOP.

      rs_result-type            = gc_success.
      rs_result-contract_number = lv_contract_number.
      IF lv_all_messages IS INITIAL.
        rs_result-validation_step = 'BAPI_SUCCESS'.
        rs_result-message         = |Contract { lv_contract_number } created successfully|.
      ELSE.
        rs_result-validation_step = 'BAPI_WARNING'.
        rs_result-message         = |Contract { lv_contract_number } created with warnings: { lv_all_messages }|.
      ENDIF.

      IF p_test IS INITIAL.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING wait = 'X'
          IMPORTING return = ls_commit_return.
        IF ls_commit_return-type = gc_error OR ls_commit_return-type = gc_abort.
          rs_result-type = gc_error.
          rs_result-validation_step = 'COMMIT_ERROR'.
          rs_result-message = |Contract created but commit failed: { ls_commit_return-message }|.
        ENDIF.
      ELSE.
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
        rs_result-validation_step = 'TEST_MODE'.
        rs_result-message = |Test Mode: Contract { lv_contract_number } would be created (rolled back)|.
      ENDIF.
    ELSE.
      rs_result-type = gc_error.
      rs_result-validation_step = 'BAPI_NO_CONTRACT'.
      rs_result-message = 'BAPI call completed but no contract number was generated'.
      RETURN.
    ENDIF.
  ENDMETHOD.

  METHOD call_bapi_change.
    DATA: ls_headdatain         TYPE bapicchead,
          ls_headdatainx        TYPE bapiccheadx,
          lt_return             TYPE TABLE OF bapiret2,
          ls_return             TYPE bapiret2,
          ls_commit_return      TYPE bapiret2,
          lv_all_messages       TYPE string,
          lv_separator          TYPE string VALUE '; ',
          lv_ccnum_alpha        TYPE bapicckey-condition_contract_number,
          lv_customer_formatted TYPE kna1-kunnr,
          lv_any_change         TYPE abap_bool VALUE abap_false.

    rs_result-row_number      = iv_row_number.
    rs_result-contract_type   = is_data-contract_type.
    rs_result-process_variant = is_data-process_variant.
    rs_result-ext_num         = is_data-ext_num.
    rs_result-cust_owner      = is_data-cust_owner.
    rs_result-date_from       = is_data-date_from.
    rs_result-date_to         = is_data-date_to.
    rs_result-contract_number = is_data-contract_number.

    IF is_data-contract_number IS INITIAL.
      rs_result-type = gc_error.
      rs_result-validation_step = 'INPUT_VALIDATION'.
      rs_result-message = 'Condition Contract Number is mandatory for change'.
      RETURN.
    ENDIF.

    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING input  = is_data-contract_number
      IMPORTING output = lv_ccnum_alpha.

    IF is_data-cust_owner IS NOT INITIAL.
      lv_customer_formatted = format_customer( is_data-cust_owner ).
    ENDIF.

    " Headdata updates (flag only fields intended to change)
    IF is_data-process_variant IS NOT INITIAL.
      ls_headdatain-process_variant = is_data-process_variant.
      ls_headdatainx-process_variant = 'X'.
      lv_any_change = abap_true.
    ENDIF.

    IF is_data-date_from IS NOT INITIAL.
      ls_headdatain-validity_date_from = is_data-date_from.
      ls_headdatainx-validity_date_from = 'X'.
      lv_any_change = abap_true.
    ENDIF.

    IF is_data-date_to IS NOT INITIAL.
      ls_headdatain-validity_date_to = is_data-date_to.
      ls_headdatainx-validity_date_to = 'X'.
      lv_any_change = abap_true.
    ENDIF.

    IF lv_customer_formatted IS NOT INITIAL.
      ls_headdatain-customer_owner = lv_customer_formatted.
      ls_headdatainx-customer_owner = 'X'.
      lv_any_change = abap_true.
    ENDIF.

    IF is_data-currency IS NOT INITIAL.
      FIELD-SYMBOLS: <fs_cur_val>  TYPE any,
                     <fs_cur_iso>  TYPE any,
                     <fs_curx_val> TYPE any,
                     <fs_curx_iso> TYPE any.

      ASSIGN COMPONENT 'CURRENCY'     OF STRUCTURE ls_headdatain  TO <fs_cur_val>.
      IF <fs_cur_val> IS ASSIGNED.
        <fs_cur_val> = is_data-currency.
        lv_any_change = abap_true.
      ENDIF.

      ASSIGN COMPONENT 'CURRENCY_ISO' OF STRUCTURE ls_headdatain  TO <fs_cur_iso>.
      IF <fs_cur_iso> IS ASSIGNED.
        <fs_cur_iso> = is_data-currency.
        lv_any_change = abap_true.
      ENDIF.

      ASSIGN COMPONENT 'CURRENCY'     OF STRUCTURE ls_headdatainx TO <fs_curx_val>.
      IF <fs_curx_val> IS ASSIGNED.
        <fs_curx_val> = 'X'.
      ENDIF.

      ASSIGN COMPONENT 'CURRENCY_ISO' OF STRUCTURE ls_headdatainx TO <fs_curx_iso>.
      IF <fs_curx_iso> IS ASSIGNED.
        <fs_curx_iso> = 'X'.
      ENDIF.
    ENDIF.

    DATA: lt_bapiccbvb  TYPE TABLE OF bapiccbvb,
          lt_bapiccbvbx TYPE TABLE OF bapiccbvbx.

    IF lv_customer_formatted IS NOT INITIAL.
      DATA(ls_bapiccbvb)  = VALUE bapiccbvb(
        order_key        = '1'
        include_exclude  = 'I'
        fieldcombination = '0001'
        customer_new     = lv_customer_formatted
        customer_key     = '1' ).
      APPEND ls_bapiccbvb TO lt_bapiccbvb.

      DATA(ls_bapiccbvbx) = VALUE bapiccbvbx(
        order_key   = '1'
        updateflag  = 'U' ).
      APPEND ls_bapiccbvbx TO lt_bapiccbvbx.
      lv_any_change = abap_true.
    ENDIF.

    IF lv_any_change IS INITIAL.
      rs_result-type = gc_warning.
      rs_result-validation_step = 'NO_CHANGE'.
      rs_result-message = |No change requested: no updatable fields provided for contract { lv_ccnum_alpha }|.
      RETURN.
    ENDIF.

    CALL FUNCTION 'BAPI_CONDITION_CONTRACT_CHANGE'
      EXPORTING
        conditioncontractnumber = lv_ccnum_alpha
        headdatain              = ls_headdatain
        headdatainx             = ls_headdatainx
      TABLES
        bvbdatain               = lt_bapiccbvb
        bvbdatainx              = lt_bapiccbvbx
        return                  = lt_return.

    DATA(lv_detail) = VALUE string( ).
    CLEAR lv_all_messages.
    LOOP AT lt_return INTO ls_return WHERE type = gc_abort.
      lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
      IF ls_return-message_v1 IS NOT INITIAL OR
         ls_return-message_v2 IS NOT INITIAL OR
         ls_return-message_v3 IS NOT INITIAL OR
         ls_return-message_v4 IS NOT INITIAL.
        lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
      ENDIF.
      IF lv_all_messages IS INITIAL.
        lv_all_messages = |Abort: { lv_detail }|.
      ELSE.
        lv_all_messages = |{ lv_all_messages }{ lv_separator }Abort: { lv_detail }|.
      ENDIF.
    ENDLOOP.

    IF lv_all_messages IS NOT INITIAL.
      rs_result-type = gc_error.
      rs_result-validation_step = 'BAPI_ABORT'.
      rs_result-message = lv_all_messages.
      RETURN.
    ENDIF.

    CLEAR lv_all_messages.
    LOOP AT lt_return INTO ls_return WHERE type = gc_error.
      lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
      IF ls_return-message_v1 IS NOT INITIAL OR
         ls_return-message_v2 IS NOT INITIAL OR
         ls_return-message_v3 IS NOT INITIAL OR
         ls_return-message_v4 IS NOT INITIAL.
        lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
      ENDIF.
      IF lv_all_messages IS INITIAL.
        lv_all_messages = |Error: { lv_detail }|.
      ELSE.
        lv_all_messages = |{ lv_all_messages }{ lv_separator }Error: { lv_detail }|.
      ENDIF.
    ENDLOOP.

    IF lv_all_messages IS NOT INITIAL.
      rs_result-type = gc_error.
      rs_result-validation_step = 'BAPI_ERROR'.
      rs_result-message = lv_all_messages.
      RETURN.
    ENDIF.

    DATA(lv_has_success) = abap_false.
    READ TABLE lt_return INTO ls_return WITH KEY type = gc_success.
    IF sy-subrc = 0.
      lv_has_success = abap_true.
    ENDIF.

    CLEAR lv_all_messages.
    LOOP AT lt_return INTO ls_return WHERE type = gc_warning.
      lv_detail = |{ ls_return-id }-{ ls_return-number } { ls_return-message }|.
      IF ls_return-message_v1 IS NOT INITIAL OR
         ls_return-message_v2 IS NOT INITIAL OR
         ls_return-message_v3 IS NOT INITIAL OR
         ls_return-message_v4 IS NOT INITIAL.
        lv_detail = |{ lv_detail } ( { ls_return-message_v1 } { ls_return-message_v2 } { ls_return-message_v3 } { ls_return-message_v4 } )|.
      ENDIF.
      IF lv_all_messages IS INITIAL.
        lv_all_messages = |Warning: { lv_detail }|.
      ELSE.
        lv_all_messages = |{ lv_all_messages }{ lv_separator }Warning: { lv_detail }|.
      ENDIF.
    ENDLOOP.

    IF lv_has_success = abap_true.
      rs_result-type            = gc_success.
      rs_result-contract_number = lv_ccnum_alpha.
      rs_result-validation_step = COND string( WHEN lv_all_messages IS INITIAL THEN 'BAPI_SUCCESS' ELSE 'BAPI_WARNING' ).
      rs_result-message         = COND string( WHEN lv_all_messages IS INITIAL
                                               THEN |Contract { lv_ccnum_alpha } changed successfully|
                                               ELSE |Contract { lv_ccnum_alpha } changed with warnings: { lv_all_messages }| ).
    ELSE.
      rs_result-type            = gc_warning.
      rs_result-contract_number = lv_ccnum_alpha.
      rs_result-validation_step = COND string( WHEN lv_all_messages IS INITIAL THEN 'BAPI_NO_ERROR' ELSE 'BAPI_NO_ERROR_WARNING' ).
      rs_result-message         = COND string( WHEN lv_all_messages IS INITIAL
                                               THEN |No explicit success message returned by BAPI; verify contract { lv_ccnum_alpha }|
                                               ELSE |No success message; warnings: { lv_all_messages }| ).
    ENDIF.

    IF p_test IS INITIAL AND rs_result-type <> gc_error.
      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          wait   = 'X'
        IMPORTING
          return = ls_commit_return.

      IF ls_commit_return-type = gc_error OR ls_commit_return-type = gc_abort.
        rs_result-type = gc_error.
        rs_result-validation_step = 'COMMIT_ERROR'.
        rs_result-message = |Change saved but commit failed: { ls_commit_return-message }|.
      ENDIF.
    ELSEIF p_test IS NOT INITIAL AND rs_result-type <> gc_error.
      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      rs_result-validation_step = 'TEST_MODE'.
      rs_result-message = |Test Mode: Contract { lv_ccnum_alpha } would be changed (rolled back)|.
    ENDIF.
  ENDMETHOD.

  METHOD create_condition_contracts.
    CLEAR: gt_message, gv_success_count, gv_error_count.
    DATA: lv_row_number        TYPE i,
          ls_validation_result TYPE lty_validation_result.

    IF lines( gt_condition_data ) = 0.
      DATA(ls_no_data_message) = VALUE lty_message(
        row_number = 0
        type = gc_error
        validation_step = 'NO_DATA'
        message = 'No data found to process. Please check your Excel file.' ).
      APPEND ls_no_data_message TO gt_message.
      ADD 1 TO gv_error_count.
      RETURN.
    ENDIF.

    FIELD-SYMBOLS: <fs_header> TYPE lty_condition_data.
    DATA: ls_header     TYPE lty_condition_data,
          lt_bvb_group  TYPE STANDARD TABLE OF lty_condition_data WITH EMPTY KEY.
    DATA lv_has_header TYPE abap_bool VALUE abap_false.
    DATA lv_header_row TYPE i VALUE 0.

    LOOP AT gt_condition_data ASSIGNING FIELD-SYMBOL(<ls_row>).
      ADD 1 TO lv_row_number.

      DATA(lv_is_header_row) = xsdbool( <ls_row>-contract_type IS NOT INITIAL OR
                                        <ls_row>-process_variant IS NOT INITIAL OR
                                        <ls_row>-cust_owner IS NOT INITIAL OR
                                        <ls_row>-date_from IS NOT INITIAL OR
                                        <ls_row>-date_to   IS NOT INITIAL OR
                                        <ls_row>-vkorg     IS NOT INITIAL OR
                                        <ls_row>-vtweg     IS NOT INITIAL OR
                                        <ls_row>-spart     IS NOT INITIAL ).

      IF lv_is_header_row = abap_true.
        IF lv_has_header = abap_true.
          TRY.
              ls_validation_result = validate_data_enhanced( is_data = ls_header iv_row_number = lv_header_row ).
              IF ls_validation_result-valid = abap_false.
                DATA(ls_err_hdr) = VALUE lty_message(
                  row_number = lv_header_row
                  contract_type = ls_header-contract_type
                  process_variant = ls_header-process_variant
                  ext_num = ls_header-ext_num
                  cust_owner = ls_header-cust_owner
                  date_from = ls_header-date_from
                  date_to = ls_header-date_to
                  type = gc_error
                  validation_step = ls_validation_result-validation_step
                  message = ls_validation_result-error_message ).
                APPEND ls_err_hdr TO gt_message.
                ADD 1 TO gv_error_count.
              ELSE.
                DATA(ls_group_result) = call_bapi_create_group( is_header = ls_header it_bvb_rows = lt_bvb_group iv_header_row = lv_header_row ).
                APPEND ls_group_result TO gt_message.
                IF ls_group_result-type = gc_success.
                  ADD 1 TO gv_success_count.
                ELSE.
                  ADD 1 TO gv_error_count.
                ENDIF.
              ENDIF.
            CATCH cx_root INTO DATA(lo_grp_ex).
              DATA(ls_grp_err) = VALUE lty_message(
                row_number = lv_header_row
                contract_type = ls_header-contract_type
                process_variant = ls_header-process_variant
                ext_num = ls_header-ext_num
                cust_owner = ls_header-cust_owner
                date_from = ls_header-date_from
                date_to = ls_header-date_to
                type = gc_error
                validation_step = 'GROUP_EXCEPTION'
                message = |Group Processing Exception: { lo_grp_ex->get_text( ) }| ).
              APPEND ls_grp_err TO gt_message.
              ADD 1 TO gv_error_count.
          ENDTRY.
        ENDIF.

        CLEAR ls_header.
        ls_header-contract_type   = <ls_row>-contract_type.
        ls_header-process_variant = <ls_row>-process_variant.
        ls_header-ext_num         = <ls_row>-ext_num.
        ls_header-cust_owner      = <ls_row>-cust_owner.
        ls_header-date_from       = <ls_row>-date_from.
        ls_header-date_to         = <ls_row>-date_to.
        ls_header-vkorg           = <ls_row>-vkorg.
        ls_header-vtweg           = <ls_row>-vtweg.
        ls_header-spart           = <ls_row>-spart.
        ls_header-zterm           = <ls_row>-zterm.
        ls_header-settl_cal_accr  = <ls_row>-settl_cal_accr.
        ls_header-assignment      = <ls_row>-assignment.
        ls_header-settl_cal_part  = <ls_row>-settl_cal_part.
        ls_header-settl_cal_delta = <ls_row>-settl_cal_delta.
        ls_header-cc_curr         = <ls_row>-cc_curr.
        ls_header-exchange_rate   = <ls_row>-exchange_rate.
        ls_header-exchange_rate_type = <ls_row>-exchange_rate_type.

        lv_has_header = abap_true.
        lv_header_row = lv_row_number.
        ASSIGN ls_header TO <fs_header>.

        CLEAR lt_bvb_group.
        IF <ls_row>-order_key IS NOT INITIAL OR <ls_row>-matnr IS NOT INITIAL OR <ls_row>-kunnr IS NOT INITIAL.
          APPEND <ls_row> TO lt_bvb_group.
        ENDIF.

      ELSE.
        IF lv_has_header = abap_true.
          IF <ls_row>-order_key IS NOT INITIAL OR <ls_row>-matnr IS NOT INITIAL OR <ls_row>-kunnr IS NOT INITIAL.
            APPEND <ls_row> TO lt_bvb_group.
          ENDIF.
        ELSE.
          DATA(ls_orphan) = VALUE lty_message(
            row_number = lv_row_number
            type = gc_error
            validation_step = 'ORPHAN_BVB'
            message = 'BVB row encountered before any header row.' ).
          APPEND ls_orphan TO gt_message.
          ADD 1 TO gv_error_count.
        ENDIF.
      ENDIF.
    ENDLOOP.

    IF lv_has_header = abap_true.
      TRY.
          ls_validation_result = validate_data_enhanced( is_data = ls_header iv_row_number = lv_header_row ).
          IF ls_validation_result-valid = abap_false.
            DATA(ls_err_hdr_last) = VALUE lty_message(
              row_number = lv_header_row
              contract_type = ls_header-contract_type
              process_variant = ls_header-process_variant
              ext_num = ls_header-ext_num
              cust_owner = ls_header-cust_owner
              date_from = ls_header-date_from
              date_to = ls_header-date_to
              type = gc_error
              validation_step = ls_validation_result-validation_step
              message = ls_validation_result-error_message ).
            APPEND ls_err_hdr_last TO gt_message.
            ADD 1 TO gv_error_count.
          ELSE.
            DATA(ls_group_result_last) = call_bapi_create_group( is_header = ls_header it_bvb_rows = lt_bvb_group iv_header_row = lv_header_row ).
            APPEND ls_group_result_last TO gt_message.
            IF ls_group_result_last-type = gc_success.
              ADD 1 TO gv_success_count.
            ELSE.
              ADD 1 TO gv_error_count.
            ENDIF.
          ENDIF.
        CATCH cx_root INTO DATA(lo_grp_ex_last).
          DATA(ls_grp_err_last) = VALUE lty_message(
            row_number = lv_header_row
            contract_type = ls_header-contract_type
            process_variant = ls_header-process_variant
            ext_num = ls_header-ext_num
            cust_owner = ls_header-cust_owner
            date_from = ls_header-date_from
            date_to = ls_header-date_to
            type = gc_error
            validation_step = 'GROUP_EXCEPTION'
            message = |Group Processing Exception: { lo_grp_ex_last->get_text( ) }| ).
          APPEND ls_grp_err_last TO gt_message.
          ADD 1 TO gv_error_count.
      ENDTRY.
    ENDIF.

    IF gv_error_count = 0.
      MESSAGE |Processing completed successfully: { gv_success_count } contracts created| TYPE 'S'.
    ELSEIF gv_success_count = 0.
      MESSAGE |Processing completed with errors: { gv_error_count } contracts failed| TYPE 'S' DISPLAY LIKE 'E'.
    ELSE.
      MESSAGE |Processing completed: { gv_success_count } successful, { gv_error_count } failed| TYPE 'S' DISPLAY LIKE 'W'.
    ENDIF.
  ENDMETHOD.

  METHOD change_condition_contracts.
    CLEAR: gt_message, gv_success_count, gv_error_count.
    DATA: lv_row_number        TYPE i,
          ls_validation_result TYPE lty_validation_result.

    IF lines( gt_condition_data ) = 0.
      DATA(ls_no_data_message) = VALUE lty_message(
        row_number = 0
        type = gc_error
        validation_step = 'NO_DATA'
        message = 'No data found to process. Please check your Excel file.'
      ).
      APPEND ls_no_data_message TO gt_message.
      ADD 1 TO gv_error_count.
      RETURN.
    ENDIF.

    LOOP AT gt_condition_data ASSIGNING FIELD-SYMBOL(<ls_data>).
      ADD 1 TO lv_row_number.

      TRY.
          ls_validation_result = validate_data_change( is_data = <ls_data> iv_row_number = lv_row_number ).
          IF ls_validation_result-valid = abap_false.
            DATA(ls_missing_key) = VALUE lty_message(
              row_number = lv_row_number
              contract_number = <ls_data>-contract_number
              contract_type = <ls_data>-contract_type
              process_variant = <ls_data>-process_variant
              ext_num = <ls_data>-ext_num
              cust_owner = <ls_data>-cust_owner
              date_from = <ls_data>-date_from
              date_to = <ls_data>-date_to
              type = gc_error
              validation_step = ls_validation_result-validation_step
              message = ls_validation_result-error_message
            ).
            APPEND ls_missing_key TO gt_message.
            ADD 1 TO gv_error_count.
            CONTINUE.
          ENDIF.

          TRY.
              DATA(ls_bapi_result) = call_bapi_change( is_data = <ls_data> iv_row_number = lv_row_number ).
              APPEND ls_bapi_result TO gt_message.
              IF ls_bapi_result-type = gc_success.
                ADD 1 TO gv_success_count.
              ELSEIF ls_bapi_result-type = gc_error.
                ADD 1 TO gv_error_count.
              ENDIF.
            CATCH cx_root INTO DATA(lo_bapi_exception).
              DATA(ls_bapi_error) = VALUE lty_message(
                row_number = lv_row_number
                contract_number = <ls_data>-contract_number
                contract_type = <ls_data>-contract_type
                process_variant = <ls_data>-process_variant
                ext_num = <ls_data>-ext_num
                cust_owner = <ls_data>-cust_owner
                date_from = <ls_data>-date_from
                date_to = <ls_data>-date_to
                type = gc_error
                validation_step = 'BAPI_EXCEPTION'
                message = |BAPI Exception: { lo_bapi_exception->get_text( ) }|
              ).
              APPEND ls_bapi_error TO gt_message.
              ADD 1 TO gv_error_count.
          ENDTRY.

        CATCH cx_root INTO DATA(lo_general_exception).
          DATA(ls_general_error) = VALUE lty_message(
            row_number = lv_row_number
            contract_number = <ls_data>-contract_number
            contract_type = <ls_data>-contract_type
            process_variant = <ls_data>-process_variant
            ext_num = <ls_data>-ext_num
            cust_owner = <ls_data>-cust_owner
            date_from = <ls_data>-date_from
            date_to = <ls_data>-date_to
            type = gc_error
            validation_step = 'PROCESSING_EXCEPTION'
            message = |Processing Exception: { lo_general_exception->get_text( ) }|
          ).
          APPEND ls_general_error TO gt_message.
          ADD 1 TO gv_error_count.
      ENDTRY.
    ENDLOOP.

    IF gv_error_count = 0.
      MESSAGE |Processing completed successfully: { gv_success_count } contracts changed| TYPE 'S'.
    ELSEIF gv_success_count = 0.
      MESSAGE |Processing completed with errors: { gv_error_count } changes failed| TYPE 'S' DISPLAY LIKE 'E'.
    ELSE.
      MESSAGE |Processing completed: { gv_success_count } successful, { gv_error_count } failed| TYPE 'S' DISPLAY LIKE 'W'.
    ENDIF.
  ENDMETHOD.

  METHOD get_data.
    CLEAR: gt_condition_data.
    DATA: lt_data_tab TYPE solix_tab.

    IF NOT p_file IS INITIAL.
      cl_gui_frontend_services=>gui_upload(
        EXPORTING
          filename                = CONV #( p_file )
          filetype                = 'BIN'
        IMPORTING
          filelength              = DATA(lv_filelength)
        CHANGING
          data_tab                = lt_data_tab
        EXCEPTIONS
          OTHERS                  = 1 ).

      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 INTO DATA(lv_message).
        MESSAGE lv_message TYPE 'I' DISPLAY LIKE 'E'.
        LEAVE LIST-PROCESSING.
      ELSE.
        IF NOT lt_data_tab IS INITIAL.
          cl_bcs_convert=>solix_to_xstring(
            EXPORTING
              it_solix   = lt_data_tab
              iv_size    = lv_filelength
            RECEIVING
              ev_xstring = DATA(lv_xstring) ).

          DATA(lo_excel) = NEW cl_fdt_xl_spreadsheet(
                               document_name = CONV #( p_file )
                               xdocument     = lv_xstring ).

          lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
            IMPORTING
              worksheet_names = DATA(lt_worksheets) ).

          DATA(lv_worksheet) = VALUE #( lt_worksheets[ 1 ] OPTIONAL ).

          DATA(lo_data_ref) = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_worksheet ).
          ASSIGN lo_data_ref->* TO FIELD-SYMBOL(<lfs_data_tab>).
          IF <lfs_data_tab> IS NOT ASSIGNED.
            MESSAGE 'Unable to read data from Excel worksheet.' TYPE 'I' DISPLAY LIKE 'E'.
            LEAVE LIST-PROCESSING.
          ENDIF.

          IF lines( <lfs_data_tab> ) LE 1.
            MESSAGE 'Excel file contains no data rows (only headers found).' TYPE 'I' DISPLAY LIKE 'E'.
            LEAVE LIST-PROCESSING.
          ENDIF.

          DATA(lv_data_rows) = 0.
          IF p_change IS INITIAL.
            " CREATE: map header and BVB fields as per template
            LOOP AT <lfs_data_tab> ASSIGNING FIELD-SYMBOL(<lfs_data_row>).
              IF sy-tabix = 1. CONTINUE. ENDIF.
              APPEND INITIAL LINE TO gt_condition_data ASSIGNING FIELD-SYMBOL(<lfs_condition_data>).
              DO 47 TIMES.
                ASSIGN COMPONENT sy-index OF STRUCTURE <lfs_data_row> TO FIELD-SYMBOL(<lfs_value>).
                IF <lfs_value> IS ASSIGNED.
                  CASE sy-index.
                    WHEN 1.  <lfs_condition_data>-contract_type      = <lfs_value>.
                    WHEN 2.  <lfs_condition_data>-process_variant    = <lfs_value>.
                    WHEN 3.  <lfs_condition_data>-ext_num            = <lfs_value>.
                    WHEN 4.  <lfs_condition_data>-cust_owner         = <lfs_value>.
                    WHEN 5.  <lfs_condition_data>-date_from          = convert_excel_date( <lfs_value> ).
                    WHEN 6.  <lfs_condition_data>-date_to            = convert_excel_date( <lfs_value> ).
                    WHEN 7.  <lfs_condition_data>-vkorg              = <lfs_value>.
                    WHEN 8.  <lfs_condition_data>-vtweg              = <lfs_value>.
                    WHEN 9.  <lfs_condition_data>-spart              = <lfs_value>.
                    WHEN 10. <lfs_condition_data>-zterm              = <lfs_value>.
                    WHEN 11. <lfs_condition_data>-settl_cal_accr     = <lfs_value>.
                    WHEN 12. <lfs_condition_data>-assignment         = <lfs_value>.
                    WHEN 13. <lfs_condition_data>-settl_cal_part     = <lfs_value>.
                    WHEN 14. <lfs_condition_data>-settl_cal_delta    = <lfs_value>.
                    WHEN 15. <lfs_condition_data>-cc_curr            = <lfs_value>.
                    WHEN 16. <lfs_condition_data>-exchange_rate      = <lfs_value>.
                    WHEN 17. <lfs_condition_data>-exchange_rate_type = <lfs_value>.
                    WHEN 18. <lfs_condition_data>-order_key          = <lfs_value>.
                    WHEN 19. <lfs_condition_data>-fieldcomb          = <lfs_value>.
                    WHEN 20. <lfs_condition_data>-incl_excl          = <lfs_value>.
                    WHEN 21. <lfs_condition_data>-kunnr              = <lfs_value>.
                    WHEN 22. <lfs_condition_data>-matnr              = <lfs_value>.
                    WHEN 23. <lfs_condition_data>-vkorg_bvb          = <lfs_value>.
                    WHEN 24. <lfs_condition_data>-augru              = <lfs_value>.
                    WHEN 25. <lfs_condition_data>-kunhier            = <lfs_value>.
                    WHEN 26. <lfs_condition_data>-prodh              = <lfs_value>.
                    WHEN 27. <lfs_condition_data>-mvgr1              = <lfs_value>.
                    WHEN 28. <lfs_condition_data>-mvgr2              = <lfs_value>.
                    WHEN 29. <lfs_condition_data>-mvgr3              = <lfs_value>.
                    WHEN 30. <lfs_condition_data>-mvgr4              = <lfs_value>.
                    WHEN 31. <lfs_condition_data>-mvgr5              = <lfs_value>.
                    WHEN 32. <lfs_condition_data>-kvgr1              = <lfs_value>.
                    WHEN 33. <lfs_condition_data>-kvgr2              = <lfs_value>.
                    WHEN 34. <lfs_condition_data>-kvgr3              = <lfs_value>.
                    WHEN 35. <lfs_condition_data>-kvgr4              = <lfs_value>.
                    WHEN 36. <lfs_condition_data>-kvgr5              = <lfs_value>.
                    WHEN 37. <lfs_condition_data>-prodh1             = <lfs_value>.
                    WHEN 38. <lfs_condition_data>-prodh2             = <lfs_value>.
                    WHEN 39. <lfs_condition_data>-prodh3             = <lfs_value>.
                    WHEN 40. <lfs_condition_data>-zzprodh4           = <lfs_value>.
                    WHEN 41. <lfs_condition_data>-zzprodh5           = <lfs_value>.
                    WHEN 42. <lfs_condition_data>-zzprodh6           = <lfs_value>.
                    WHEN 43. <lfs_condition_data>-zzprodh7           = <lfs_value>.
                    WHEN 44. <lfs_condition_data>-vtweg_bvb          = <lfs_value>.
                    WHEN 45. <lfs_condition_data>-werks              = <lfs_value>.
                    WHEN 46. <lfs_condition_data>-spart_bvb          = <lfs_value>.
                    WHEN 47. <lfs_condition_data>-kunrg              = <lfs_value>.
                  ENDCASE.
                ENDIF.
              ENDDO.
              ADD 1 TO lv_data_rows.
            ENDLOOP.
          ELSE.
            LOOP AT <lfs_data_tab> ASSIGNING FIELD-SYMBOL(<lfs_data_row_c>).
              IF sy-tabix = 1. CONTINUE. ENDIF.
              APPEND INITIAL LINE TO gt_condition_data ASSIGNING FIELD-SYMBOL(<lfs_condition_data_c>).
              DO 8 TIMES.
                ASSIGN COMPONENT sy-index OF STRUCTURE <lfs_data_row_c> TO FIELD-SYMBOL(<lfs_value_c>).
                IF <lfs_value_c> IS ASSIGNED.
                  CASE sy-index.
                    WHEN 1. <lfs_condition_data_c>-contract_number = <lfs_value_c>.
                    WHEN 2. <lfs_condition_data_c>-currency        = <lfs_value_c>.
                    WHEN 3. <lfs_condition_data_c>-contract_type   = <lfs_value_c>.
                    WHEN 4. <lfs_condition_data_c>-process_variant = <lfs_value_c>.
                    WHEN 5. <lfs_condition_data_c>-ext_num         = <lfs_value_c>.
                    WHEN 6. <lfs_condition_data_c>-cust_owner      = <lfs_value_c>.
                    WHEN 7. <lfs_condition_data_c>-date_from       = convert_excel_date( <lfs_value_c> ).
                    WHEN 8. <lfs_condition_data_c>-date_to         = convert_excel_date( <lfs_value_c> ).
                  ENDCASE.
                ENDIF.
              ENDDO.
              ADD 1 TO lv_data_rows.
            ENDLOOP.
          ENDIF.

          IF lv_data_rows = 0.
            MESSAGE 'Excel file contains no data rows (only headers found).' TYPE 'I' DISPLAY LIKE 'E'.
            LEAVE LIST-PROCESSING.
          ENDIF.

          MESSAGE |Successfully read { lv_data_rows } records from Excel file| TYPE 'S'.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD display_results.
    DATA: lo_columns TYPE REF TO cl_salv_columns_table,
          lo_column  TYPE REF TO cl_salv_column_table.

    TRY.
        CALL METHOD cl_salv_table=>factory
          IMPORTING
            r_salv_table = DATA(lo_obj_alv)
          CHANGING
            t_table      = gt_message.
      CATCH cx_salv_msg INTO DATA(lo_salv_msg).
        MESSAGE lo_salv_msg->get_text( ) TYPE 'I' DISPLAY LIKE 'E'.
        RETURN.
    ENDTRY.

    lo_columns = lo_obj_alv->get_columns( ).
    lo_columns->set_optimize( ).

    TRY.
        lo_column ?= lo_columns->get_column( 'ROW_NUMBER' ).
        lo_column->set_medium_text( 'Row #' ).
        lo_column->set_short_text( 'Row' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'CONTRACT_NUMBER' ).
        lo_column->set_medium_text( 'Contract Number' ).
        lo_column->set_short_text( 'Contract' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'CONTRACT_TYPE' ).
        lo_column->set_medium_text( 'Contract Type' ).
        lo_column->set_short_text( 'ContrType' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'PROCESS_VARIANT' ).
        lo_column->set_medium_text( 'Process Variant' ).
        lo_column->set_short_text( 'ProcVar' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'EXT_NUM' ).
        lo_column->set_medium_text( 'External ID' ).
        lo_column->set_short_text( 'ExtID' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'CUST_OWNER' ).
        lo_column->set_medium_text( 'Customer' ).
        lo_column->set_short_text( 'Customer' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'DATE_FROM' ).
        lo_column->set_medium_text( 'Valid From' ).
        lo_column->set_short_text( 'From' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'DATE_TO' ).
        lo_column->set_medium_text( 'Valid To' ).
        lo_column->set_short_text( 'To' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'TYPE' ).
        lo_column->set_medium_text( 'Status' ).
        lo_column->set_short_text( 'Status' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'VALIDATION_STEP' ).
        lo_column->set_medium_text( 'Validation Step' ).
        lo_column->set_short_text( 'Step' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= lo_columns->get_column( 'MESSAGE' ).
        lo_column->set_medium_text( 'Message' ).
        lo_column->set_short_text( 'Message' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    DATA(lo_functions) = lo_obj_alv->get_functions( ).
    lo_functions->set_default( abap_true ).

    DATA(lo_display) = lo_obj_alv->get_display_settings( ).
    lo_display->set_striped_pattern( abap_true ).

    DATA(lo_header) = NEW cl_salv_form_layout_grid( ).
    DATA(lo_h_label) = lo_header->create_label( row = 1 column = 1 ).
    lo_h_label->set_text( |Condition Contract Processing Results| ).

    DATA(lo_h_info) = lo_header->create_label( row = 2 column = 1 ).
    lo_h_info->set_text( |Total Records: { lines( gt_condition_data ) } - Success: { gv_success_count } - Errors: { gv_error_count }| ).

    IF p_test IS NOT INITIAL.
      DATA(lo_h_test) = lo_header->create_label( row = 3 column = 1 ).
      lo_h_test->set_text( |TEST MODE - No actual contracts were created/changed| ).
    ENDIF.

    lo_obj_alv->set_top_of_list( lo_header ).

    lo_obj_alv->display( ).
  ENDMETHOD.

  METHOD download_template.
    DATA: lt_template_data TYPE STANDARD TABLE OF lty_condition_data WITH DEFAULT KEY.

    TRY.
        cl_salv_table=>factory(
          EXPORTING
            container_name = space
          IMPORTING
            r_salv_table   = DATA(lo_alv)
          CHANGING
            t_table        = lt_template_data
        ).
      CATCH cx_salv_msg INTO DATA(lo_salv_msg).
        MESSAGE lo_salv_msg->get_text( ) TYPE 'I' DISPLAY LIKE 'E'.
        LEAVE LIST-PROCESSING.
    ENDTRY.

    CALL METHOD set_alv_column_headers
      EXPORTING
        io_alv = lo_alv.

    DATA(lo_columns) = lo_alv->get_columns( ).

    TRY.
        DATA(lo_col_ccnum) = lo_columns->get_column( 'CONTRACT_NUMBER' ).
        IF p_change IS INITIAL.
          lo_col_ccnum->set_technical( abap_true ).
        ELSE.
          lo_col_ccnum->set_long_text( 'Condition Contract Number' ).
          lo_col_ccnum->set_medium_text( 'Cond Contract Number' ).
          lo_col_ccnum->set_short_text( 'Contract' ).
        ENDIF.
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        DATA(lo_col_curr) = lo_columns->get_column( 'CURRENCY' ).
        IF p_change IS INITIAL.
          lo_col_curr->set_technical( abap_true ).
        ELSE.
          lo_col_curr->set_long_text( 'Contract Currency (e.g., USD)' ).
          lo_col_curr->set_medium_text( 'Contract Currency' ).
          lo_col_curr->set_short_text( 'Currency' ).
        ENDIF.
      CATCH cx_salv_not_found.
    ENDTRY.

    DATA(lv_xstring) = VALUE xstring( ).
    IF lo_alv IS BOUND.
      lo_alv->to_xml(
        EXPORTING
          xml_type    = if_salv_bs_xml=>c_type_xlsx
        RECEIVING
          xml         = lv_xstring
      ).
    ENDIF.

    IF lv_xstring IS NOT INITIAL.
      cl_bcs_convert=>xstring_to_solix(
        EXPORTING
          iv_xstring = lv_xstring
        RECEIVING
          et_solix   = DATA(lt_solix)
      ).
    ENDIF.

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize            = xstrlen( lv_xstring )
        filename                = CONV #( p_file )
        filetype                = 'BIN'
      CHANGING
        data_tab                = lt_solix
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
        not_supported_by_gui    = 22
        error_no_gui            = 23
        OTHERS                  = 24
    ).

    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ELSE.
      MESSAGE |Template downloaded successfully to: { p_file }| TYPE 'S'.
    ENDIF.
  ENDMETHOD.

  METHOD check_input_file.
    DATA: lv_selected_folder TYPE string,
          lt_filetable       TYPE filetable,
          lv_rc              TYPE i,
          lv_uaction         TYPE i.

    DATA(lv_title) = COND string( WHEN p_down IS NOT INITIAL THEN 'Select Download Directory'
                                  ELSE 'Select Excel File to Upload' ).

    IF p_down IS NOT INITIAL.
      CALL METHOD cl_gui_frontend_services=>directory_browse(
        EXPORTING
          window_title         = lv_title
        CHANGING
          selected_folder      = lv_selected_folder
        EXCEPTIONS
          cntl_error           = 1
          error_no_gui         = 2
          not_supported_by_gui = 3
          OTHERS               = 4
      ).

      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ELSE.
        DATA(lv_len) = strlen( lv_selected_folder ) - 1.
        IF lv_selected_folder+lv_len(1) <> '\\' AND lv_selected_folder+lv_len(1) <> '/'.
          lv_selected_folder = lv_selected_folder && '\\'.
        ENDIF.
        DATA(lv_fname) = COND string( WHEN p_change IS INITIAL
                                       THEN 'condition_contract_template.xlsx'
                                       ELSE 'condition_contract_change_template.xlsx' ).
        p_file = |{ lv_selected_folder }{ lv_fname }|.
      ENDIF.
    ELSE.
      cl_gui_frontend_services=>file_open_dialog(
        EXPORTING
          window_title            = lv_title
          default_extension       = cl_gui_frontend_services=>filetype_excel
          file_filter             = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
        CHANGING
          file_table              = lt_filetable
          rc                      = lv_rc
          user_action             = lv_uaction
        EXCEPTIONS
          file_open_dialog_failed = 1
          cntl_error              = 2
          error_no_gui            = 3
          not_supported_by_gui    = 4
          OTHERS                  = 5 ).

      IF sy-subrc = 0 AND lv_uaction = 0 AND lines( lt_filetable ) > 0.
        p_file = VALUE #( lt_filetable[ 1 ]-filename OPTIONAL ).
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD check_file_exist.
    IF p_file IS NOT INITIAL.
      DATA(lv_path_segments) = |{ p_file }|.
      SPLIT lv_path_segments AT '\\' INTO TABLE DATA(lt_segments).
      IF lines( lt_segments ) = 0.
        SPLIT lv_path_segments AT '/' INTO TABLE lt_segments.
      ENDIF.

      DATA(lv_filename) = VALUE #( lt_segments[ lines( lt_segments ) ] OPTIONAL ).

      IF p_upload IS NOT INITIAL.
        IF lv_filename IS NOT INITIAL.
          SPLIT lv_filename AT '.' INTO TABLE DATA(lt_extension).
          DATA(lv_extension) = VALUE #( lt_extension[ lines( lt_extension ) ] OPTIONAL ).

          IF to_upper( lv_extension ) NE 'XLSX'.
            MESSAGE 'Only Excel files (.xlsx) are supported.' TYPE 'I' DISPLAY LIKE 'E'.
            LEAVE LIST-PROCESSING.
          ENDIF.
        ENDIF.

        CALL METHOD cl_gui_frontend_services=>file_exist
          EXPORTING
            file                 = p_file
          RECEIVING
            result               = DATA(lv_result)
          EXCEPTIONS
            cntl_error           = 1
            error_no_gui         = 2
            wrong_parameter      = 3
            not_supported_by_gui = 4
            OTHERS               = 5.

        IF sy-subrc <> 0 OR lv_result <> abap_true.
          MESSAGE 'Selected file does not exist.' TYPE 'I' DISPLAY LIKE 'E'.
          LEAVE LIST-PROCESSING.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD selection_screen.
    LOOP AT SCREEN.
      IF screen-name = 'P_TEST'.
        IF p_upload IS NOT INITIAL.
          screen-input = '1'.
        ELSEIF p_upload IS INITIAL.
          screen-input = '0'.
        ENDIF.
        MODIFY SCREEN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD set_alv_column_headers.
    DATA: lo_columns TYPE REF TO cl_salv_columns_table,
          lo_column  TYPE REF TO cl_salv_column_table.

    TRY.
        lo_columns = io_alv->get_columns( ).
        lo_columns->set_optimize( ).

        TRY.
            lo_column ?= lo_columns->get_column( 'CONTRACT_NUMBER' ).
            lo_column->set_long_text( 'Condition Contract Number' ).
            lo_column->set_medium_text( 'Contract Number' ).
            lo_column->set_short_text( 'Contract' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'CURRENCY' ).
            lo_column->set_long_text( 'Contract Currency (e.g., USD)' ).
            lo_column->set_medium_text( 'Contract Currency' ).
            lo_column->set_short_text( 'Currency' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'CONTRACT_TYPE' ).
            lo_column->set_long_text( 'Contract Type (e.g., ZCC1)' ).
            lo_column->set_medium_text( 'Contract Type' ).
            lo_column->set_short_text( 'ContType' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'PROCESS_VARIANT' ).
            lo_column->set_long_text( 'Process Variant (e.g., V001)' ).
            lo_column->set_medium_text( 'Process Variant' ).
            lo_column->set_short_text( 'ProcVar' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'EXT_NUM' ).
            lo_column->set_long_text( 'External Number (Optional)' ).
            lo_column->set_medium_text( 'External ID' ).
            lo_column->set_short_text( 'ExtID' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'CUST_OWNER' ).
            lo_column->set_long_text( 'Customer Number (e.g., 1000)' ).
            lo_column->set_medium_text( 'Customer' ).
            lo_column->set_short_text( 'Customer' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'DATE_FROM' ).
            lo_column->set_long_text( 'Valid From Date (YYYYMMDD or Excel date)' ).
            lo_column->set_medium_text( 'From Date' ).
            lo_column->set_short_text( 'From' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'DATE_TO' ).
            lo_column->set_long_text( 'Valid To Date (YYYYMMDD or Excel date)' ).
            lo_column->set_medium_text( 'To Date' ).
            lo_column->set_short_text( 'To' ).
          CATCH cx_salv_not_found.
        ENDTRY.
        " BVB field headers
        TRY.
            lo_column ?= lo_columns->get_column( 'ORDER_KEY' ).
            lo_column->set_long_text( 'Order Key' ).
            lo_column->set_medium_text( 'Order Key' ).
            lo_column->set_short_text( 'OrderKey' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'FIELDCOMB' ).
            lo_column->set_long_text( 'Field Combination' ).
            lo_column->set_medium_text( 'Field Comb' ).
            lo_column->set_short_text( 'FldComb' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'INCL_EXCL' ).
            lo_column->set_long_text( 'Include/Exclude (I/E)' ).
            lo_column->set_medium_text( 'Incl/Excl' ).
            lo_column->set_short_text( 'I/E' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KUNNR' ).
            lo_column->set_long_text( 'Customer' ).
            lo_column->set_medium_text( 'Customer' ).
            lo_column->set_short_text( 'Customer' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MATNR' ).
            lo_column->set_long_text( 'Material' ).
            lo_column->set_medium_text( 'Material' ).
            lo_column->set_short_text( 'Material' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'VKORG_BVB' ).
            lo_column->set_long_text( 'Sales Org (BVB)' ).
            lo_column->set_medium_text( 'Sales Org BVB' ).
            lo_column->set_short_text( 'SalesOrg' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'AUGRU' ).
            lo_column->set_long_text( 'Order Reason' ).
            lo_column->set_medium_text( 'Order Reason' ).
            lo_column->set_short_text( 'OrdRsn' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KUNHIER' ).
            lo_column->set_long_text( 'Customer Hierarchy' ).
            lo_column->set_medium_text( 'Cust Hierarchy' ).
            lo_column->set_short_text( 'CustHier' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'PRODH' ).
            lo_column->set_long_text( 'Product Hierarchy' ).
            lo_column->set_medium_text( 'Prod Hierarchy' ).
            lo_column->set_short_text( 'ProdHier' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MVGR1' ).
            lo_column->set_long_text( 'Material Group 1' ).
            lo_column->set_medium_text( 'Mat Group 1' ).
            lo_column->set_short_text( 'MatGrp1' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MVGR2' ).
            lo_column->set_long_text( 'Material Group 2' ).
            lo_column->set_medium_text( 'Mat Group 2' ).
            lo_column->set_short_text( 'MatGrp2' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MVGR3' ).
            lo_column->set_long_text( 'Material Group 3' ).
            lo_column->set_medium_text( 'Mat Group 3' ).
            lo_column->set_short_text( 'MatGrp3' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MVGR4' ).
            lo_column->set_long_text( 'Material Group 4' ).
            lo_column->set_medium_text( 'Mat Group 4' ).
            lo_column->set_short_text( 'MatGrp4' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'MVGR5' ).
            lo_column->set_long_text( 'Material Group 5' ).
            lo_column->set_medium_text( 'Mat Group 5' ).
            lo_column->set_short_text( 'MatGrp5' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KVGR1' ).
            lo_column->set_long_text( 'Customer Group 1' ).
            lo_column->set_medium_text( 'Cust Group 1' ).
            lo_column->set_short_text( 'CusGrp1' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KVGR2' ).
            lo_column->set_long_text( 'Customer Group 2' ).
            lo_column->set_medium_text( 'Cust Group 2' ).
            lo_column->set_short_text( 'CusGrp2' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KVGR3' ).
            lo_column->set_long_text( 'Customer Group 3' ).
            lo_column->set_medium_text( 'Cust Group 3' ).
            lo_column->set_short_text( 'CusGrp3' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KVGR4' ).
            lo_column->set_long_text( 'Customer Group 4' ).
            lo_column->set_medium_text( 'Cust Group 4' ).
            lo_column->set_short_text( 'CusGrp4' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KVGR5' ).
            lo_column->set_long_text( 'Customer Group 5' ).
            lo_column->set_medium_text( 'Cust Group 5' ).
            lo_column->set_short_text( 'CusGrp5' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'PRODH1' ).
            lo_column->set_long_text( 'Product Line' ).
            lo_column->set_medium_text( 'Product Line' ).
            lo_column->set_short_text( 'ProdLin' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'PRODH2' ).
            lo_column->set_long_text( 'Product Group' ).
            lo_column->set_medium_text( 'Product Group' ).
            lo_column->set_short_text( 'ProdGrp' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'PRODH3' ).
            lo_column->set_long_text( 'Brand' ).
            lo_column->set_medium_text( 'Brand' ).
            lo_column->set_short_text( 'Brand' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'ZZPRODH4' ).
            lo_column->set_long_text( 'Sub Brand' ).
            lo_column->set_medium_text( 'Sub Brand' ).
            lo_column->set_short_text( 'SubBrand' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'ZZPRODH5' ).
            lo_column->set_long_text( 'Container Size' ).
            lo_column->set_medium_text( 'Container Size' ).
            lo_column->set_short_text( 'ContSize' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'ZZPRODH6' ).
            lo_column->set_long_text( 'Count' ).
            lo_column->set_medium_text( 'Count' ).
            lo_column->set_short_text( 'Count' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'ZZPRODH7' ).
            lo_column->set_long_text( 'Additional Product Info' ).
            lo_column->set_medium_text( 'Add Prod Info' ).
            lo_column->set_short_text( 'AddProd' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'VTWEG_BVB' ).
            lo_column->set_long_text( 'Distribution Channel (BVB)' ).
            lo_column->set_medium_text( 'Dist Chan BVB' ).
            lo_column->set_short_text( 'DistChBV' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'WERKS' ).
            lo_column->set_long_text( 'Plant' ).
            lo_column->set_medium_text( 'Plant' ).
            lo_column->set_short_text( 'Plant' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'SPART_BVB' ).
            lo_column->set_long_text( 'Division (BVB)' ).
            lo_column->set_medium_text( 'Division BVB' ).
            lo_column->set_short_text( 'DivBVB' ).
          CATCH cx_salv_not_found.
        ENDTRY.

        TRY.
            lo_column ?= lo_columns->get_column( 'KUNRG' ).
            lo_column->set_long_text( 'Payer' ).
            lo_column->set_medium_text( 'Payer' ).
            lo_column->set_short_text( 'Payer' ).
          CATCH cx_salv_not_found.
        ENDTRY.
  ENDMETHOD.

  METHOD has_data.
    rv_has_data = xsdbool( lines( gt_condition_data ) > 0 ).
  ENDMETHOD.

ENDCLASS.


AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  DATA(lo_class) = lcl_condition_contract=>get_instance( ).
  CALL METHOD lo_class->check_input_file.

AT SELECTION-SCREEN OUTPUT.
  DATA(lo_class) = lcl_condition_contract=>get_instance( ).
  CALL METHOD lo_class->selection_screen.

START-OF-SELECTION.
  DATA(lo_class) = lcl_condition_contract=>get_instance( ).

  CALL METHOD lo_class->authorization_check.
  CALL METHOD lo_class->check_file_exist.

  IF p_create = 'X' AND p_down = 'X'.
    CALL METHOD lo_class->download_template.

  ELSEIF p_create = 'X' AND p_upload = 'X'.
    CALL METHOD lo_class->get_data.
    IF lo_class->has_data( ) = abap_true.
      CALL METHOD lo_class->create_condition_contracts.
      CALL METHOD lo_class->display_results.
    ELSE.
      MESSAGE 'No data found in Excel or format is invalid. Please check your Excel file structure.' TYPE 'I'.
    ENDIF.

  ELSEIF p_change = 'X' AND p_down = 'X'.
    CALL METHOD lo_class->download_template.

  ELSEIF p_change = 'X' AND p_upload = 'X'.
    CALL METHOD lo_class->get_data.
    IF lo_class->has_data( ) = abap_true.
      CALL METHOD lo_class->change_condition_contracts.
      CALL METHOD lo_class->display_results.
    ELSE.
      MESSAGE 'No data found in Excel or format is invalid. Please check your Excel file structure.' TYPE 'I'.
    ENDIF.

  ELSE.
    MESSAGE 'Please select Create/Change with Download or Upload and specify the file path.' TYPE 'E'.
  ENDIF.

