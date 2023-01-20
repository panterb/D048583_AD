extend view entity  ZAD_R_ORDERINGTP with 

association [0..1] to zad_i_ordering      as _zz_ExtensionZHC         on  $projection.Key1 = _zz_ExtensionZHC.Key1
composition [0..1] of zad_r_extnodetp     as _zz_ExtNodeZHC


{

    _zz_ExtNodeZHC,
    _Extension.zz_char_field_zhc as zz_char_field_zhc,
    _Extension.zz_curr_field1_zhc as zz_curr_field1_zhc,
    _Extension.zz_curr_field2_zhc as zz_curr_field2_zhc,
    
    _zz_ExtensionZHC
}

