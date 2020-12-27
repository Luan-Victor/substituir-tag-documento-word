*&---------------------------------------------------------------------*
*& Report  ZWORDMAILMERGE_EXAMPLE
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zwordmailmerge_example.

TYPES: BEGIN OF ty_mard,
        matnr TYPE mard-matnr,
        werks TYPE mard-werks,
        lgort TYPE mard-lgort,
        umlme TYPE mard-umlme,
       END OF ty_mard.

DATA: lo_word      TYPE REF TO zcl_word_mailmerge,
      lt_word_data TYPE ztword_data,
      ls_mard      TYPE ty_mard,
      lv_filename  TYPE rlgrap-filename VALUE 'C:\WordMailMerge\testdocument.docx',
      lv_path      TYPE rlgrap-filename VALUE 'C:\WordMailMerge\'.

FIELD-SYMBOLS: <word_data> LIKE LINE OF lt_word_data.

* Busca registro
SELECT SINGLE matnr werks lgort umlme
  FROM mard
  INTO ls_mard
  WHERE umlme > 0.

* Salva dados
  APPEND INITIAL LINE TO lt_word_data ASSIGNING <word_data>.
  <word_data>-field = 'MATNR'.
  <word_data>-value = ls_mard-matnr.

  APPEND INITIAL LINE TO lt_word_data ASSIGNING <word_data>.
  <word_data>-field = 'WERKS'.
  <word_data>-value = ls_mard-werks.

  APPEND INITIAL LINE TO lt_word_data ASSIGNING <word_data>.
  <word_data>-field = 'LGORT'.
  <word_data>-value = ls_mard-lgort.

  APPEND INITIAL LINE TO lt_word_data ASSIGNING <word_data>.
  <word_data>-field = 'UMLME'.
  <word_data>-value = ls_mard-umlme.

* Cria objeto que substitui campo pelos valores
  CREATE OBJECT lo_word
    EXPORTING
      i_word_data       = lt_word_data  " Table with rows data
      i_template_file   = lv_filename   " File local para Upload ou Download
      i_download_path   = lv_path       " File local para Upload ou Download
    EXCEPTIONS
      file_do_not_exist = 1
      OTHERS            = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

* Imprime documento
  lo_word->print_document( ).
