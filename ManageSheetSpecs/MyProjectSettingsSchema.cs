using System;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ManageSheetSpecs
{
  public static class MyProjectSettingsSchema
  {
    readonly static Guid schemaGuid = new Guid( 
      "{9DBE0174-AA01-4CDD-BA86-96DE1FDCE041}" );

    public static Schema GetSchema()
    {
      Schema schema = Schema.Lookup( schemaGuid );

      if( schema != null ) return schema;

      SchemaBuilder schemaBuilder =
          new SchemaBuilder( schemaGuid );

      schemaBuilder.SetSchemaName( 
        "MyProjectSettings" );

      schemaBuilder.AddSimpleField( 
        "Parameter1", typeof( int ) );

      schemaBuilder.AddSimpleField( 
        "Parameter2", typeof( string ) );

      return schemaBuilder.Finish();
    }
  }
}
