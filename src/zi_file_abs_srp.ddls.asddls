@EndUserText.label: 'Upload file action'
@Metadata.allowExtensions: true
define abstract entity ZI_FILE_ABS_SRP
  //  with parameters parameter_name : parameter_type
{

//  @Semantics.mimeType: true
  mimeType      : abap.string(0);
  fileName      : abap.string(0);
//  @Semantics.largeObject: {
//      mimeType: 'mimeType',
//      fileName: 'fileName',
//      contentDispositionPreference: #INLINE
////      acceptableMimeTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
//    }
  fileContent   : abap.rawstring(0);
  fileExtension : abap.string(0);

}
