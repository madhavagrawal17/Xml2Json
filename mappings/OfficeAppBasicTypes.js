var OfficeAppBasicTypes_Module_Factory = function () {
  var OfficeAppBasicTypes = {
    name: 'OfficeAppBasicTypes',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
    typeInfos: [{
        localName: 'Methods',
        propertyInfos: [{
            name: 'method',
            required: true,
            collection: true,
            elementName: 'Method',
            typeInfo: '.VersionedRequirement'
          }, {
            name: 'defaultMinVersion',
            attributeName: {
              localPart: 'DefaultMinVersion'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ShortLocaleAwareSettingWithId',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: 'Override',
            typeInfo: '.ShortLocaleOverride'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLResources',
        propertyInfos: [{
            name: 'url',
            minOccurs: 0,
            collection: true,
            elementName: 'Url',
            typeInfo: '.URLLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'LongStringResources',
        propertyInfos: [{
            name: 'string',
            minOccurs: 0,
            collection: true,
            elementName: 'String',
            typeInfo: '.LongLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'ShortResourceReference',
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'ShortStringResources',
        propertyInfos: [{
            name: 'string',
            minOccurs: 0,
            collection: true,
            elementName: 'String',
            typeInfo: '.ShortLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'Requirements',
        propertyInfos: [{
            name: 'sets',
            required: true,
            elementName: 'Sets',
            typeInfo: '.Sets'
          }]
      }, {
        localName: 'ResourceReference',
        propertyInfos: [{
            name: 'resid',
            required: true,
            attributeName: {
              localPart: 'resid'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'LongLocaleOverride',
        propertyInfos: [{
            name: 'locale',
            required: true,
            attributeName: {
              localPart: 'Locale'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'Value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLResourceReference',
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'MobileIconList',
        propertyInfos: [{
            name: 'image',
            required: true,
            minOccurs: 9,
            collection: true,
            elementName: 'Image',
            typeInfo: '.MobileImageResourceReference'
          }]
      }, {
        localName: 'ImageResources',
        propertyInfos: [{
            name: 'image',
            minOccurs: 0,
            collection: true,
            elementName: 'Image',
            typeInfo: '.ImageLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'URLLocaleAwareSettingWithId',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: 'Override',
            typeInfo: '.URLLocaleOverride'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLLocaleOverride',
        propertyInfos: [{
            name: 'locale',
            required: true,
            attributeName: {
              localPart: 'Locale'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'Value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'IconList',
        propertyInfos: [{
            name: 'image',
            required: true,
            collection: true,
            elementName: 'Image',
            typeInfo: '.ImageResourceReference'
          }]
      }, {
        localName: 'ImageResourceReference',
        baseTypeInfo: '.ResourceReference',
        propertyInfos: [{
            name: 'size',
            required: true,
            typeInfo: 'Integer',
            attributeName: {
              localPart: 'size'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'MobileImageResourceReference',
        baseTypeInfo: '.ResourceReference',
        propertyInfos: [{
            name: 'size',
            required: true,
            typeInfo: 'Integer',
            attributeName: {
              localPart: 'size'
            },
            type: 'attribute'
          }, {
            name: 'scale',
            required: true,
            typeInfo: 'Integer',
            attributeName: {
              localPart: 'scale'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ShortLocaleOverride',
        propertyInfos: [{
            name: 'locale',
            required: true,
            attributeName: {
              localPart: 'Locale'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'Value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Sets',
        propertyInfos: [{
            name: 'set',
            required: true,
            collection: true,
            elementName: 'Set',
            typeInfo: '.VersionedRequirement'
          }, {
            name: 'defaultMinVersion',
            attributeName: {
              localPart: 'DefaultMinVersion'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ImageLocaleAwareSettingWithId',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: 'Override',
            typeInfo: '.URLLocaleOverride'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'LongLocaleAwareSettingWithId',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: 'Override',
            typeInfo: '.LongLocaleOverride'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'LongResourceReference',
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'Resources',
        propertyInfos: [{
            name: 'images',
            elementName: 'Images',
            typeInfo: '.ImageResources'
          }, {
            name: 'urls',
            elementName: 'Urls',
            typeInfo: '.URLResources'
          }, {
            name: 'shortStrings',
            elementName: 'ShortStrings',
            typeInfo: '.ShortStringResources'
          }, {
            name: 'longStrings',
            elementName: 'LongStrings',
            typeInfo: '.LongStringResources'
          }]
      }, {
        localName: 'VersionedRequirement',
        propertyInfos: [{
            name: 'minVersion',
            attributeName: {
              localPart: 'MinVersion'
            },
            type: 'attribute'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'Name'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: []
  };
  return {
    OfficeAppBasicTypes: OfficeAppBasicTypes
  };
};
if (typeof define === 'function' && define.amd) {
  define([], OfficeAppBasicTypes_Module_Factory);
}
else {
  var OfficeAppBasicTypes_Module = OfficeAppBasicTypes_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.OfficeAppBasicTypes = OfficeAppBasicTypes_Module.OfficeAppBasicTypes;
  }
  else {
    var OfficeAppBasicTypes = OfficeAppBasicTypes_Module.OfficeAppBasicTypes;
  }
}