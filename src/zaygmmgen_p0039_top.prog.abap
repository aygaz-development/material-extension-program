*&---------------------------------------------------------------------*
*& Include zaygmmgen_p0039_top
*&---------------------------------------------------------------------*

CLASS lcl_program DEFINITION DEFERRED.

TYPE-POOLS icon.

DATA gs_manuel_create_scr TYPE zaygmm_s001.

TYPES tt_matnr_r TYPE RANGE OF matnr.

TYPES: BEGIN OF ty_plant_copy,
         werks_source TYPE werks_d,
         werks_target TYPE werks_d,
         matnr        TYPE matnr,
         matnr_r      TYPE tt_matnr_r,
       END OF ty_plant_copy.

DATA gs_plant_copy_scr TYPE ty_plant_copy.
