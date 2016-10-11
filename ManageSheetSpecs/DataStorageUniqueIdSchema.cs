using System;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ManageSheetSpecs
{
  static class DataStorageUniqueIdSchema
  {
    static readonly Guid schemaGuid = new Guid( 
      "{EEEFD606-7262-4782-93F0-2DA87D5AE6E4}" );

    public static Schema GetSchema()
    {
      Schema schema = Schema.Lookup( schemaGuid );

      if( schema != null )
        return schema;

      SchemaBuilder schemaBuilder = new SchemaBuilder( 
        schemaGuid );

      schemaBuilder.SetSchemaName( 
        "DataStorageUniqueId" );

      schemaBuilder.AddSimpleField( 
        "Id", typeof( Guid ) );

      return schemaBuilder.Finish();
    }
  }
}
