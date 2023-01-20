CLASS zad_cl_read_appl_excel_email DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .



  PUBLIC SECTION.
      class-DATA: report TYPE zad_xco_report.
      INTERFACES if_oo_adt_classrun .
      INTERFACES if_apj_dt_exec_object .
      INTERFACES if_apj_rt_exec_object .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zad_cl_read_appl_excel_email IMPLEMENTATION.

 METHOD if_apj_dt_exec_object~get_parameters.
    " Parameter Description for Application Jobs Template
    et_parameter_def = VALUE #(
        ( selname = 'P_DESCR' kind = if_apj_dt_exec_object=>parameter     datatype = 'C' length = 80 param_text = 'Description'   lowercase_ind = abap_true changeable_ind = abap_true )
      ).

    " Parameter Table for Application Jobs Template
    et_parameter_val = VALUE #(
      ( selname = 'P_DESCR' kind = if_apj_dt_exec_object=>parameter     sign = 'I' option = 'EQ' low = 'Job Template for Online Shop' )
    ).
  ENDMETHOD.

    METHOD if_apj_rt_exec_object~execute.

    ENDMETHOD.

  METHOD if_oo_adt_classrun~main.
     TRY.

    " Generate the XLSX report.
      DATA(lo_document) = xco_cp_xlsx=>document->empty(
        )->write_access( ).
      DATA(lo_worksheet) = lo_document->get_workbook(
        )->worksheet->at_position( 1 ).

       " Read custom BO data
       "DATA lt_online_shop TYPE STANDARD TABLE OF zad_i_online_shop WITH DEFAULT KEY.
       SELECT FROM zad_online_shop fields order_id, ordereditem, creationdate
       into TABLE @data(lt_online_shop).

       " Write the header of the worksheet
        DATA(lo_cursor) = lo_worksheet->cursor(
         io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
         io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
         ).
        lo_cursor->get_cell( )->value->write_from( 'Custom BO Application Data' ).
        lo_cursor->move_down( ).

         " Prepare the data that shall be written into the XLSX report.
          DATA(lv_date) = CONV d( xco_cp=>sy->date( )->as( xco_cp_time=>format->abap )->value ).
          DATA(lv_time) = CONV t( xco_cp=>sy->time( )->as( xco_cp_time=>format->abap )->value ).

         "Write the header information
             "Write the current date
             lo_cursor->move_down( )->get_cell( )->value->write_from( 'Date:' ).
             lo_cursor->move_right( )->get_cell( )->value->write_from( lv_date ).

             " Write the current time
             lo_cursor->move_down( )->move_left( )->get_cell( )->value->write_from( 'Time:' ).
             lo_cursor->move_right( )->get_cell( )->value->write_from( lv_time ).

             lo_cursor->move_down( )->move_down( ).

            " Write the application data into the worksheet
             DATA(lo_pattern) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                 )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
                 )->from_row( lo_cursor->position->row
                 )->get_pattern( ).

             lo_worksheet->select( lo_pattern
              )->row_stream(
              )->operation->write_from( REF #( lt_online_shop )
              )->set_value_transformation( xco_cp_xlsx_write_access=>value_transformation->best_effort
              )->execute( ).

      DATA(lv_file_content) = lo_document->get_file_content( ).

      zad_cl_read_appl_excel_email=>report = VALUE #(
        Excel         = lv_file_content
        ExcelMimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ExcelFilename = 'report.xlsx'
      ).

        DATA(lo_mail) = cl_bcs_mail_message=>create_instance( ).

        lo_mail->set_sender( 'noreply@sap.com' ).
        lo_mail->add_recipient( 'alda.dollani@sap.com' ).
        lo_mail->set_subject( 'Notification: Job complete' ).

        lo_mail->set_main( cl_bcs_mail_textpart=>create_instance(
          iv_content      = '<h1>Confirmation</h1><p>The job to process the items uploaded via XLSX is finished successfully.</p>'
          iv_content_type = 'text/html'
        ) ).

        lo_mail->add_attachment( cl_bcs_mail_binarypart=>create_instance(
         iv_content      = lv_file_content
         iv_content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
         iv_filename     = 'ApplicationData.xlsx'
        ) ).

        lo_mail->send( IMPORTING et_status = DATA(lt_status) ).

       CATCH cx_bcs_mail INTO DATA(lx_mail).
        out->write( EXPORTING data = |Exception: { lx_mail->get_longtext( ) }| ).
       CATCH cx_bali_runtime INTO DATA(lx_bali_runtime).

    endtry.

    ENDMETHOD.

ENDCLASS.
