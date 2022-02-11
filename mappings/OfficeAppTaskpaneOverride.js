var OfficeAppTaskpaneOverride_Module_Factory = function () {
  var OfficeAppTaskpaneOverride = {
    name: 'OfficeAppTaskpaneOverride',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
    typeInfos: [{
        localName: 'VersionOverridesV10',
        typeName: 'VersionOverridesV1_0',
        propertyInfos: [{
            name: 'description',
            elementName: 'Description',
            typeInfo: '.LongResourceReference'
          }, {
            name: 'requirements',
            elementName: 'Requirements',
            typeInfo: '.Requirements'
          }, {
            name: 'hosts',
            elementName: 'Hosts',
            typeInfo: '.Hosts'
          }, {
            name: 'resources',
            elementName: 'Resources',
            typeInfo: '.Resources'
          }, {
            name: 'webApplicationInfo',
            elementName: 'WebApplicationInfo',
            typeInfo: '.WebApplicationInfo'
          }, {
            name: 'equivalentAddins',
            elementName: 'EquivalentAddins',
            typeInfo: '.EquivalentAddins'
          }, {
            name: 'connectedServiceControls',
            elementName: 'ConnectedServiceControls',
            typeInfo: '.ConnectedServiceControls'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'WebAppAuthorization',
        propertyInfos: [{
            name: 'resource',
            required: true,
            elementName: 'Resource'
          }, {
            name: 'scopes',
            required: true,
            elementName: 'Scopes',
            typeInfo: '.WebApplicationScopes'
          }]
      }, {
        localName: 'WebApplicationInfo',
        propertyInfos: [{
            name: 'id',
            required: true,
            elementName: 'Id'
          }, {
            name: 'msaId',
            elementName: 'MsaId'
          }, {
            name: 'resource',
            required: true,
            elementName: 'Resource'
          }, {
            name: 'scopes',
            required: true,
            elementName: 'Scopes',
            typeInfo: '.WebApplicationScopes'
          }, {
            name: 'authorizations',
            elementName: 'Authorizations',
            typeInfo: '.WebAppAuthorizations'
          }]
      }, {
        localName: 'EquivalentAddins',
        propertyInfos: [{
            name: 'equivalentAddin',
            required: true,
            collection: true,
            elementName: 'EquivalentAddin',
            typeInfo: '.EquivalentAddin'
          }]
      }, {
        localName: 'ShortLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ShortLocaleOverride'
        },
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
        localName: 'ShortLocaleAwareSettingWithId',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ShortLocaleAwareSettingWithId'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
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
        localName: 'Resources',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'Resources'
        },
        propertyInfos: [{
            name: 'images',
            elementName: {
              localPart: 'Images',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.ImageResources'
          }, {
            name: 'urls',
            elementName: {
              localPart: 'Urls',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.URLResources'
          }, {
            name: 'shortStrings',
            elementName: {
              localPart: 'ShortStrings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.ShortStringResources'
          }, {
            name: 'longStrings',
            elementName: {
              localPart: 'LongStrings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.LongStringResources'
          }]
      }, {
        localName: 'OfficeMenu',
        propertyInfos: [{
            name: 'control',
            required: true,
            collection: true,
            elementName: 'Control',
            typeInfo: '.UIControl'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'URLLocaleOverride'
        },
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
        localName: 'Metadata',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'MenuItems',
        propertyInfos: [{
            name: 'item',
            required: true,
            collection: true,
            elementName: 'Item',
            typeInfo: '.MenuItem'
          }]
      }, {
        localName: 'URLLocaleAwareSettingWithId',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'URLLocaleAwareSettingWithId'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
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
        localName: 'MenuItem',
        baseTypeInfo: '.UIControlWithOptionalIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: 'Action',
            typeInfo: '.Action'
          }, {
            name: 'enabled',
            elementName: 'Enabled',
            typeInfo: 'Boolean'
          }]
      }, {
        localName: 'ExtensionPoint'
      }, {
        localName: 'UIControlWithIcon',
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            required: true,
            elementName: 'Icon',
            typeInfo: '.IconList'
          }]
      }, {
        localName: 'CustomFunctions',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'script',
            required: true,
            elementName: 'Script',
            typeInfo: '.Script'
          }, {
            name: 'page',
            required: true,
            elementName: 'Page',
            typeInfo: '.Page'
          }, {
            name: 'metadata',
            required: true,
            elementName: 'Metadata',
            typeInfo: '.Metadata'
          }, {
            name: 'namespace',
            elementName: 'Namespace',
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'ContextMenu',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'officeMenu',
            required: true,
            collection: true,
            elementName: 'OfficeMenu',
            typeInfo: '.OfficeMenu'
          }]
      }, {
        localName: 'OfficeTab',
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'group',
            required: true,
            collection: true,
            elementName: 'Group',
            typeInfo: '.Group'
          }]
      }, {
        localName: 'IconList',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'IconList'
        },
        propertyInfos: [{
            name: 'image',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Image',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.ImageResourceReference'
          }]
      }, {
        localName: 'Button',
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: 'Action',
            typeInfo: '.Action'
          }, {
            name: 'enabled',
            elementName: 'Enabled',
            typeInfo: 'Boolean'
          }]
      }, {
        localName: 'Methods',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'Methods'
        },
        propertyInfos: [{
            name: 'method',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Method',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.VersionedRequirement'
          }, {
            name: 'defaultMinVersion',
            attributeName: {
              localPart: 'DefaultMinVersion'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Document',
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: 'DesktopFormFactor',
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'Supertip',
        propertyInfos: [{
            name: 'title',
            required: true,
            elementName: 'Title',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'description',
            required: true,
            elementName: 'Description',
            typeInfo: '.LongResourceReference'
          }]
      }, {
        localName: 'Notebook',
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: 'DesktopFormFactor',
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'WebApplicationScopes',
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: 'Scope'
          }]
      }, {
        localName: 'FormFactor',
        propertyInfos: [{
            name: 'getStarted',
            elementName: 'GetStarted',
            typeInfo: '.GetStarted'
          }, {
            name: 'functionFile',
            elementName: 'FunctionFile',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'extensionPoint',
            required: true,
            collection: true,
            elementName: 'ExtensionPoint',
            typeInfo: '.ExtensionPoint'
          }]
      }, {
        localName: 'ShowTaskpane',
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'taskpaneId',
            elementName: 'TaskpaneId'
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'title',
            elementName: 'Title',
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'LongLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'LongLocaleOverride'
        },
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
        localName: 'GetStarted',
        propertyInfos: [{
            name: 'title',
            required: true,
            elementName: 'Title',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'description',
            required: true,
            elementName: 'Description',
            typeInfo: '.LongResourceReference'
          }, {
            name: 'learnMoreUrl',
            required: true,
            elementName: 'LearnMoreUrl',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'VersionedRequirement',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'VersionedRequirement'
        },
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
      }, {
        localName: 'OfficeGroup',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Host',
        propertyInfos: [{
            name: 'runtimes',
            elementName: 'Runtimes',
            typeInfo: '.Runtimes'
          }, {
            name: 'allFormFactors',
            elementName: 'AllFormFactors',
            typeInfo: '.AllFormFactors'
          }]
      }, {
        localName: 'EquivalentAddin',
        propertyInfos: [{
            name: 'progId',
            elementName: 'ProgId'
          }, {
            name: 'displayName',
            elementName: 'DisplayName'
          }, {
            name: 'fileName',
            elementName: 'FileName'
          }, {
            name: 'type',
            required: true,
            elementName: 'Type'
          }]
      }, {
        localName: 'Action'
      }, {
        localName: 'Group',
        propertyInfos: [{
            name: 'overriddenByRibbonApi',
            elementName: 'OverriddenByRibbonApi',
            typeInfo: 'Boolean'
          }, {
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'icon',
            required: true,
            elementName: 'Icon',
            typeInfo: '.IconList'
          }, {
            name: 'controlOrOfficeControl',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: 'Control',
                typeInfo: '.UIControl'
              }, {
                elementName: 'OfficeControl',
                typeInfo: '.OfficeControl'
              }],
            type: 'elements'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'PrimaryCommandSurface',
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'WebAppAuthorizations',
        propertyInfos: [{
            name: 'authorization',
            required: true,
            collection: true,
            elementName: 'Authorization',
            typeInfo: '.WebAppAuthorization'
          }]
      }, {
        localName: 'OfficeControl',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ExecuteFunction',
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'functionName',
            required: true,
            elementName: 'FunctionName'
          }]
      }, {
        localName: 'ConnectedServiceControlsScopes',
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: 'Scope'
          }]
      }, {
        localName: 'Menu',
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'items',
            required: true,
            elementName: 'Items',
            typeInfo: '.MenuItems'
          }]
      }, {
        localName: 'ImageLocaleAwareSettingWithId',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ImageLocaleAwareSettingWithId'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
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
        localName: 'ShortStringResources',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ShortStringResources'
        },
        propertyInfos: [{
            name: 'string',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'String',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.ShortLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'Runtimes',
        propertyInfos: [{
            name: 'runtime',
            required: true,
            collection: true,
            elementName: 'Runtime',
            typeInfo: '.Runtime'
          }]
      }, {
        localName: 'URLResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'URLResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'ConnectedServiceControls',
        propertyInfos: [{
            name: 'scopes',
            required: true,
            elementName: 'Scopes',
            typeInfo: '.ConnectedServiceControlsScopes'
          }]
      }, {
        localName: 'CommandSurfaceExtensionPoint',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'officeTab',
            minOccurs: 0,
            collection: true,
            elementName: 'OfficeTab',
            typeInfo: '.OfficeTab'
          }, {
            name: 'customTab',
            minOccurs: 0,
            collection: true,
            elementName: 'CustomTab',
            typeInfo: '.CustomTab'
          }]
      }, {
        localName: 'Workbook',
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: 'DesktopFormFactor',
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'Runtime',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: 'Override',
            typeInfo: '.RuntimeOverride'
          }, {
            name: 'resid',
            required: true,
            attributeName: {
              localPart: 'resid'
            },
            type: 'attribute'
          }, {
            name: 'lifetime',
            attributeName: {
              localPart: 'lifetime'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Tab',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'LongStringResources',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'LongStringResources'
        },
        propertyInfos: [{
            name: 'string',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'String',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.LongLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'LongResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'LongResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'AllFormFactors',
        propertyInfos: [{
            name: 'extensionPoint',
            elementName: 'ExtensionPoint',
            typeInfo: '.CustomFunctions'
          }]
      }, {
        localName: 'Presentation',
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: 'DesktopFormFactor',
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'MobileImageResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'MobileImageResourceReference'
        },
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
        localName: 'Sets',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'Sets'
        },
        propertyInfos: [{
            name: 'set',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Set',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.VersionedRequirement'
          }, {
            name: 'defaultMinVersion',
            attributeName: {
              localPart: 'DefaultMinVersion'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLResources',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'URLResources'
        },
        propertyInfos: [{
            name: 'url',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Url',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.URLLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'ImageResources',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ImageResources'
        },
        propertyInfos: [{
            name: 'image',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Image',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.ImageLocaleAwareSettingWithId'
          }]
      }, {
        localName: 'MobileIconList',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'MobileIconList'
        },
        propertyInfos: [{
            name: 'image',
            required: true,
            minOccurs: 9,
            collection: true,
            elementName: {
              localPart: 'Image',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.MobileImageResourceReference'
          }]
      }, {
        localName: 'ResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ResourceReference'
        },
        propertyInfos: [{
            name: 'resid',
            required: true,
            attributeName: {
              localPart: 'resid'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Page',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'Script',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'LongLocaleAwareSettingWithId',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'LongLocaleAwareSettingWithId'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
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
        localName: 'UIControl',
        propertyInfos: [{
            name: 'overriddenByRibbonApi',
            elementName: 'OverriddenByRibbonApi',
            typeInfo: 'Boolean'
          }, {
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'supertip',
            required: true,
            elementName: 'Supertip',
            typeInfo: '.Supertip'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'RuntimeOverride',
        propertyInfos: [{
            name: 'resid',
            required: true,
            attributeName: {
              localPart: 'resid'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Hosts',
        propertyInfos: [{
            name: 'host',
            required: true,
            collection: true,
            elementName: 'Host',
            typeInfo: '.Host'
          }]
      }, {
        localName: 'Requirements',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'Requirements'
        },
        propertyInfos: [{
            name: 'sets',
            required: true,
            elementName: {
              localPart: 'Sets',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0'
            },
            typeInfo: '.Sets'
          }]
      }, {
        localName: 'ShortResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ShortResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'ImageResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ImageResourceReference'
        },
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
        localName: 'CustomTab',
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'groupOrOfficeGroup',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: 'Group',
                typeInfo: '.Group'
              }, {
                elementName: 'OfficeGroup',
                typeInfo: '.OfficeGroup'
              }],
            type: 'elements'
          }, {
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'insertBefore',
            required: true,
            elementName: 'InsertBefore'
          }, {
            name: 'insertAfter',
            required: true,
            elementName: 'InsertAfter'
          }]
      }, {
        localName: 'UIControlWithOptionalIcon',
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            elementName: 'Icon',
            typeInfo: '.IconList'
          }]
      }],
    elementInfos: [{
        typeInfo: '.VersionOverridesV10',
        elementName: 'VersionOverrides'
      }]
  };
  return {
    OfficeAppTaskpaneOverride: OfficeAppTaskpaneOverride
  };
};
if (typeof define === 'function' && define.amd) {
  define([], OfficeAppTaskpaneOverride_Module_Factory);
}
else {
  var OfficeAppTaskpaneOverride_Module = OfficeAppTaskpaneOverride_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.OfficeAppTaskpaneOverride = OfficeAppTaskpaneOverride_Module.OfficeAppTaskpaneOverride;
  }
  else {
    var OfficeAppTaskpaneOverride = OfficeAppTaskpaneOverride_Module.OfficeAppTaskpaneOverride;
  }
}