var OfficeAppManifest_Module_Factory = function () {
  var OfficeAppManifest = {
    name: 'OfficeAppManifest',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides\/1.1',
    typeInfos: [{
        localName: 'FormFactor',
        propertyInfos: [{
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
        localName: 'UIControlWithOptionalIcon',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'UIControlWithOptionalIcon'
        },
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            elementName: {
              localPart: 'Icon',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.IconList'
          }]
      }, {
        localName: 'ItemHasRegularExpressionMatch',
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'regExName',
            required: true,
            attributeName: {
              localPart: 'RegExName'
            },
            type: 'attribute'
          }, {
            name: 'regExValue',
            required: true,
            attributeName: {
              localPart: 'RegExValue'
            },
            type: 'attribute'
          }, {
            name: 'propertyName',
            required: true,
            typeInfo: '.PropertyName',
            attributeName: {
              localPart: 'PropertyName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'URLLocaleAwareSetting',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'URLLocaleAwareSetting'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleOverride'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
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
        localName: 'MobileButton',
        baseTypeInfo: '.MobileUIControlWithIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: 'Action',
            typeInfo: '.Action'
          }]
      }, {
        localName: 'Button',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Button'
        },
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: {
              localPart: 'Action',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Action'
          }]
      }, {
        localName: 'ShortLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
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
        localName: 'LocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'LocaleOverride'
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
        localName: 'AppDomains',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'AppDomains'
        },
        propertyInfos: [{
            name: 'appDomain',
            required: true,
            collection: true,
            elementName: {
              localPart: 'AppDomain',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }]
      }, {
        localName: 'Hosts',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Hosts'
        },
        propertyInfos: [{
            name: 'host',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Host',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Host'
          }]
      }, {
        localName: 'AppointmentOrganizerAutoRun',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'Host',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Host'
        },
        propertyInfos: [{
            name: 'runtimes',
            elementName: {
              localPart: 'Runtimes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Runtimes'
          }, {
            name: 'allFormFactors',
            elementName: {
              localPart: 'AllFormFactors',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.AllFormFactors'
          }]
      }, {
        localName: 'ExtendedPermissions',
        propertyInfos: [{
            name: 'extendedPermission',
            required: true,
            collection: true,
            elementName: 'ExtendedPermission'
          }]
      }, {
        localName: 'ItemHasKnownEntity',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ItemHasKnownEntity'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'entityType',
            required: true,
            attributeName: {
              localPart: 'EntityType'
            },
            type: 'attribute'
          }, {
            name: 'regExFilter',
            attributeName: {
              localPart: 'RegExFilter'
            },
            type: 'attribute'
          }, {
            name: 'filterName',
            attributeName: {
              localPart: 'FilterName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
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
        localName: 'AppointmentOrganizerCommandSurface',
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'MailHost',
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'runtimes',
            elementName: 'Runtimes',
            typeInfo: '.Runtimes'
          }, {
            name: 'desktopFormFactor',
            elementName: 'DesktopFormFactor',
            typeInfo: '.FormFactorWithSupportsSharedFolders'
          }, {
            name: 'mobileFormFactor',
            elementName: 'MobileFormFactor',
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'Supertip',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Supertip'
        },
        propertyInfos: [{
            name: 'title',
            required: true,
            elementName: {
              localPart: 'Title',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'description',
            required: true,
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.LongResourceReference'
          }]
      }, {
        localName: 'Tab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Tab'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ShortResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'ShortResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'Sets',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Sets'
        },
        propertyInfos: [{
            name: 'set',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Set',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
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
        localName: 'TargetDialects',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'TargetDialects'
        },
        propertyInfos: [{
            name: 'targetDialect',
            required: true,
            collection: true,
            elementName: {
              localPart: 'TargetDialect',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }]
      }, {
        localName: 'MobileMessageReadCommandSurface',
        baseTypeInfo: '.MobileCommandSurfaceExtensionPoint'
      }, {
        localName: 'RuleCollection',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'RuleCollection'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'rule',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Rule',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Rule'
          }, {
            name: 'mode',
            required: true,
            typeInfo: '.LogicalOperator',
            attributeName: {
              localPart: 'Mode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'RuleCollection',
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'rule',
            required: true,
            collection: true,
            elementName: 'Rule',
            typeInfo: '.Rule'
          }, {
            name: 'mode',
            required: true,
            typeInfo: '.LogicalOperator',
            attributeName: {
              localPart: 'Mode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ShortLocaleAwareSetting',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ShortLocaleAwareSetting'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ShortLocaleOverride'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'UIControlWithIcon',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'UIControlWithIcon'
        },
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            required: true,
            elementName: {
              localPart: 'Icon',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.IconList'
          }]
      }, {
        localName: 'Action',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Action'
        }
      }, {
        localName: 'ShowTaskpane',
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'supportsPinning',
            elementName: 'SupportsPinning',
            typeInfo: 'Boolean'
          }]
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
        localName: 'WebApplicationScopes',
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: 'Scope'
          }]
      }, {
        localName: 'MobileMessageComposeCommandSurface',
        baseTypeInfo: '.MobileCommandSurfaceExtensionPoint'
      }, {
        localName: 'MenuItem',
        baseTypeInfo: '.UIControlWithOptionalIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: 'Action',
            typeInfo: '.Action'
          }]
      }, {
        localName: 'WebAppAuthorization',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'WebAppAuthorization'
        },
        propertyInfos: [{
            name: 'resource',
            required: true,
            elementName: {
              localPart: 'Resource',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.WebApplicationScopes'
          }]
      }, {
        localName: 'UIControl',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'UIControl'
        },
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'tooltip',
            elementName: {
              localPart: 'Tooltip',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'supertip',
            required: true,
            elementName: {
              localPart: 'Supertip',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
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
        localName: 'WebApplicationInfo',
        propertyInfos: [{
            name: 'id',
            required: true,
            elementName: 'Id'
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
        localName: 'ConnectedServiceControls',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'ConnectedServiceControls'
        },
        propertyInfos: [{
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.ConnectedServiceControlsScopes'
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
        localName: 'MessageReadCommandSurface',
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
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
        localName: 'GetStarted',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'GetStarted'
        },
        propertyInfos: [{
            name: 'title',
            required: true,
            elementName: {
              localPart: 'Title',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'description',
            required: true,
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.LongResourceReference'
          }, {
            name: 'learnMoreUrl',
            required: true,
            elementName: {
              localPart: 'LearnMoreUrl',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'RequirementsToken',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'RequirementsToken'
        },
        baseTypeInfo: '.Token',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.RequirementsTokenOverride'
          }]
      }, {
        localName: 'FormSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'FormSettings'
        },
        propertyInfos: [{
            name: 'form',
            required: true,
            maxOccurs: 2,
            collection: true,
            elementName: {
              localPart: 'Form',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.FormType'
          }]
      }, {
        localName: 'ItemIs',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ItemIs'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'itemType',
            required: true,
            attributeName: {
              localPart: 'ItemType'
            },
            type: 'attribute'
          }, {
            name: 'itemClass',
            attributeName: {
              localPart: 'ItemClass'
            },
            type: 'attribute'
          }, {
            name: 'includeSubClasses',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IncludeSubClasses'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Menu',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Menu'
        },
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'items',
            required: true,
            elementName: {
              localPart: 'Items',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.MenuItems'
          }]
      }, {
        localName: 'UIControl',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'tooltip',
            elementName: 'Tooltip',
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
        localName: 'ConnectedServiceControls',
        propertyInfos: [{
            name: 'scopes',
            required: true,
            elementName: 'Scopes',
            typeInfo: '.ConnectedServiceControlsScopes'
          }]
      }, {
        localName: 'Requirement',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Requirement'
        },
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'Name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'FormFactor',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'FormFactor'
        },
        propertyInfos: [{
            name: 'functionFile',
            elementName: {
              localPart: 'FunctionFile',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }, {
            name: 'extensionPoint',
            required: true,
            collection: true,
            elementName: {
              localPart: 'ExtensionPoint',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ExtensionPoint'
          }]
      }, {
        localName: 'MobileUIControlWithIcon',
        baseTypeInfo: '.MobileUIControl',
        propertyInfos: [{
            name: 'icon',
            required: true,
            elementName: 'Icon',
            typeInfo: '.MobileIconList'
          }]
      }, {
        localName: 'Requirements',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Requirements'
        },
        propertyInfos: [{
            name: 'sets',
            elementName: {
              localPart: 'Sets',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Sets'
          }, {
            name: 'methods',
            elementName: {
              localPart: 'Methods',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Methods'
          }]
      }, {
        localName: 'MobileAppointmentAttendeeCommandSurface',
        baseTypeInfo: '.MobileCommandSurfaceExtensionPoint'
      }, {
        localName: 'RuleCollection',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'RuleCollection'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'rule',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Rule',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Rule'
          }, {
            name: 'mode',
            required: true,
            typeInfo: '.LogicalOperator',
            attributeName: {
              localPart: 'Mode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'LaunchEvent',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'launchEvents',
            required: true,
            elementName: 'LaunchEvents',
            typeInfo: '.LaunchEvents'
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'Presentation',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Presentation'
        },
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: {
              localPart: 'DesktopFormFactor',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'MailApp',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'MailApp'
        },
        baseTypeInfo: '.OfficeApp',
        propertyInfos: [{
            name: 'requirements',
            required: true,
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.MailAppRequirements'
          }, {
            name: 'formSettings',
            required: true,
            elementName: {
              localPart: 'FormSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.FormSettings'
          }, {
            name: 'permissions',
            elementName: {
              localPart: 'Permissions',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            values: ['Restricted', 'ReadItem', 'ReadWriteItem', 'ReadWriteMailbox']
          }, {
            name: 'rule',
            required: true,
            elementName: {
              localPart: 'Rule',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Rule'
          }, {
            name: 'disableEntityHighlighting',
            elementName: {
              localPart: 'DisableEntityHighlighting',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Boolean'
          }, {
            name: 'versionOverrides',
            elementName: {
              localPart: 'VersionOverrides',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.VersionOverridesV10'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'ItemReadTabletMailAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemReadTabletMailAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'requestedHeight',
            required: true,
            elementName: {
              localPart: 'RequestedHeight',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Int'
          }]
      }, {
        localName: 'LaunchEvents',
        propertyInfos: [{
            name: 'launchEvent',
            required: true,
            collection: true,
            elementName: 'LaunchEvent',
            typeInfo: '.LaunchEventDefinition'
          }]
      }, {
        localName: 'FormFactor',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'FormFactor'
        },
        propertyInfos: [{
            name: 'getStarted',
            elementName: {
              localPart: 'GetStarted',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.GetStarted'
          }, {
            name: 'functionFile',
            elementName: {
              localPart: 'FunctionFile',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }, {
            name: 'extensionPoint',
            required: true,
            collection: true,
            elementName: {
              localPart: 'ExtensionPoint',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ExtensionPoint'
          }]
      }, {
        localName: 'LongLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
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
        localName: 'ItemHasAttachment',
        baseTypeInfo: '.Rule'
      }, {
        localName: 'FormFactorWithSupportsSharedFolders',
        propertyInfos: [{
            name: 'supportsSharedFolders',
            elementName: 'SupportsSharedFolders',
            typeInfo: 'Boolean'
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
        localName: 'WebAppAuthorizations',
        propertyInfos: [{
            name: 'authorization',
            required: true,
            collection: true,
            elementName: 'Authorization',
            typeInfo: '.WebAppAuthorization'
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
        localName: 'TaskPaneAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'TaskPaneAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }]
      }, {
        localName: 'LaunchEventDefinition',
        propertyInfos: [{
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'Type'
            },
            type: 'attribute'
          }, {
            name: 'functionName',
            required: true,
            attributeName: {
              localPart: 'FunctionName'
            },
            type: 'attribute'
          }, {
            name: 'sendMode',
            typeInfo: '.LaunchEventSendMode',
            attributeName: {
              localPart: 'SendMode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Action'
      }, {
        localName: 'CommandSurfaceExtensionPoint',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'CommandSurfaceExtensionPoint'
        },
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'officeTab',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'OfficeTab',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.OfficeTab'
          }, {
            name: 'customTab',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'CustomTab',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.CustomTab'
          }]
      }, {
        localName: 'Notebook',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Notebook'
        },
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: {
              localPart: 'DesktopFormFactor',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.FormFactor'
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
        localName: 'VersionOverridesV10',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'VersionOverridesV1_0'
        },
        propertyInfos: [{
            name: 'webApplicationInfo',
            elementName: {
              localPart: 'WebApplicationInfo',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.WebApplicationInfo'
          }, {
            name: 'connectedServiceControls',
            elementName: {
              localPart: 'ConnectedServiceControls',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.ConnectedServiceControls'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'Group',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Group'
        },
        propertyInfos: [{
            name: 'overriddenByRibbonApi',
            elementName: {
              localPart: 'OverriddenByRibbonApi',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: 'Boolean'
          }, {
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'icon',
            required: true,
            elementName: {
              localPart: 'Icon',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.IconList'
          }, {
            name: 'controlOrOfficeControl',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: {
                  localPart: 'Control',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
                },
                typeInfo: '.UIControl'
              }, {
                elementName: {
                  localPart: 'OfficeControl',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
                },
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
        localName: 'ItemIs',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemIs'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'itemType',
            required: true,
            typeInfo: '.ItemType',
            attributeName: {
              localPart: 'ItemType'
            },
            type: 'attribute'
          }, {
            name: 'formType',
            required: true,
            typeInfo: '.ItemFormType',
            attributeName: {
              localPart: 'FormType'
            },
            type: 'attribute'
          }, {
            name: 'itemClass',
            attributeName: {
              localPart: 'ItemClass'
            },
            type: 'attribute'
          }, {
            name: 'includeSubClasses',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IncludeSubClasses'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'MobileOnlineMeetingCommandSurface',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'control',
            required: true,
            elementName: 'Control',
            typeInfo: '.MobileButton'
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
        localName: 'OfficeControl',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'OfficeControl'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
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
        localName: 'ItemEditMailAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemEditMailAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }]
      }, {
        localName: 'AppointmentOrganizerCommandSurface',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'AppointmentOrganizerCommandSurface'
        },
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'ContextMenu',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ContextMenu'
        },
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'officeMenu',
            required: true,
            collection: true,
            elementName: {
              localPart: 'OfficeMenu',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.OfficeMenu'
          }]
      }, {
        localName: 'DetectedEntity',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'requestedHeight',
            elementName: 'RequestedHeight',
            typeInfo: 'Int'
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'rule',
            required: true,
            elementName: 'Rule',
            typeInfo: '.Rule'
          }]
      }, {
        localName: 'CommandSurfaceExtensionPoint',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'CommandSurfaceExtensionPoint'
        },
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'officeTab',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'OfficeTab',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.OfficeTab'
          }, {
            name: 'customTab',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'CustomTab',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.CustomTab'
          }]
      }, {
        localName: 'ShowTaskpane',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ShowTaskpane'
        },
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'taskpaneId',
            elementName: {
              localPart: 'TaskpaneId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }, {
            name: 'title',
            elementName: {
              localPart: 'Title',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'ExtensionPoint'
      }, {
        localName: 'EquivalentAddins',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'EquivalentAddins'
        },
        propertyInfos: [{
            name: 'equivalentAddin',
            required: true,
            collection: true,
            elementName: {
              localPart: 'EquivalentAddin',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.EquivalentAddin'
          }]
      }, {
        localName: 'MobileCommandSurfaceExtensionPoint',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'group',
            required: true,
            elementName: 'Group',
            typeInfo: '.MobileGroup'
          }]
      }, {
        localName: 'Action',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Action'
        }
      }, {
        localName: 'Document',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Document'
        },
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: {
              localPart: 'DesktopFormFactor',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'OfficeTab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'OfficeTab'
        },
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'group',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Group',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Group'
          }]
      }, {
        localName: 'URLResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'URLResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'WebApplicationScopes',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'WebApplicationScopes'
        },
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Scope',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            }
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
        localName: 'Tokens',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Tokens'
        },
        propertyInfos: [{
            name: 'token',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Token',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Token'
          }]
      }, {
        localName: 'VersionOverridesV11',
        typeName: 'VersionOverridesV1_1',
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
            name: 'extendedPermissions',
            elementName: 'ExtendedPermissions',
            typeInfo: '.ExtendedPermissions'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'UIControlWithOptionalIcon',
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            elementName: 'Icon',
            typeInfo: '.IconList'
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
        localName: 'ItemHasKnownEntity',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemHasKnownEntity'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'entityType',
            required: true,
            typeInfo: '.KnownEntityType',
            attributeName: {
              localPart: 'EntityType'
            },
            type: 'attribute'
          }, {
            name: 'regExFilter',
            attributeName: {
              localPart: 'RegExFilter'
            },
            type: 'attribute'
          }, {
            name: 'filterName',
            attributeName: {
              localPart: 'FilterName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'MobileUIControl',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CustomFunctions',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'CustomFunctions'
        },
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'script',
            required: true,
            elementName: {
              localPart: 'Script',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Script'
          }, {
            name: 'page',
            required: true,
            elementName: {
              localPart: 'Page',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Page'
          }, {
            name: 'metadata',
            required: true,
            elementName: {
              localPart: 'Metadata',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Metadata'
          }, {
            name: 'namespace',
            elementName: {
              localPart: 'Namespace',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'VersionOverridesV10',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'VersionOverridesV1_0'
        },
        propertyInfos: [{
            name: 'description',
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.LongResourceReference'
          }, {
            name: 'requirements',
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Requirements'
          }, {
            name: 'hosts',
            elementName: {
              localPart: 'Hosts',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Hosts'
          }, {
            name: 'resources',
            elementName: {
              localPart: 'Resources',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Resources'
          }, {
            name: 'versionOverrides',
            elementName: 'VersionOverrides',
            typeInfo: '.VersionOverridesV11'
          }]
      }, {
        localName: 'CustomTab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'CustomTab'
        },
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'groupOrOfficeGroup',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: {
                  localPart: 'Group',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
                },
                typeInfo: '.Group'
              }, {
                elementName: {
                  localPart: 'OfficeGroup',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
                },
                typeInfo: '.OfficeGroup'
              }],
            type: 'elements'
          }, {
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'insertBefore',
            required: true,
            elementName: {
              localPart: 'InsertBefore',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'insertAfter',
            required: true,
            elementName: {
              localPart: 'InsertAfter',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
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
        localName: 'ConnectedServiceControls',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ConnectedServiceControls'
        },
        propertyInfos: [{
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ConnectedServiceControlsScopes'
          }]
      }, {
        localName: 'Runtimes',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Runtimes'
        },
        propertyInfos: [{
            name: 'runtime',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Runtime',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Runtime'
          }]
      }, {
        localName: 'WebAppAuthorization',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'WebAppAuthorization'
        },
        propertyInfos: [{
            name: 'resource',
            required: true,
            elementName: {
              localPart: 'Resource',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            }
          }, {
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.WebApplicationScopes'
          }]
      }, {
        localName: 'LocaleToken',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'LocaleToken'
        },
        baseTypeInfo: '.Token',
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.LocaleTokenOverride'
          }]
      }, {
        localName: 'MenuItem',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'MenuItem'
        },
        baseTypeInfo: '.UIControlWithOptionalIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: {
              localPart: 'Action',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Action'
          }]
      }, {
        localName: 'Hosts',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Hosts'
        },
        propertyInfos: [{
            name: 'host',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Host',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Host'
          }]
      }, {
        localName: 'WebApplicationInfo',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'WebApplicationInfo'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            elementName: {
              localPart: 'Id',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            }
          }, {
            name: 'resource',
            required: true,
            elementName: {
              localPart: 'Resource',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            }
          }, {
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.WebApplicationScopes'
          }, {
            name: 'authorizations',
            elementName: {
              localPart: 'Authorizations',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.WebAppAuthorizations'
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
        localName: 'AppointmentAttendeeCommandSurface',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'AppointmentAttendeeCommandSurface'
        },
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'UIControlWithOptionalIcon',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'UIControlWithOptionalIcon'
        },
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            elementName: {
              localPart: 'Icon',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.IconList'
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
        localName: 'Host',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Host'
        }
      }, {
        localName: 'Event',
        propertyInfos: [{
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'Type'
            },
            type: 'attribute'
          }, {
            name: 'functionExecution',
            required: true,
            typeInfo: '.EventFunctionExecutionType',
            attributeName: {
              localPart: 'FunctionExecution'
            },
            type: 'attribute'
          }, {
            name: 'functionName',
            required: true,
            attributeName: {
              localPart: 'FunctionName'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'PrimaryCommandSurface',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'PrimaryCommandSurface'
        },
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'Workbook',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Workbook'
        },
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: {
              localPart: 'DesktopFormFactor',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'LocaleAwareSetting',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'LocaleAwareSetting'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.LocaleOverride'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ItemReadDesktopMailAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemReadDesktopMailAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'requestedHeight',
            required: true,
            elementName: {
              localPart: 'RequestedHeight',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Int'
          }]
      }, {
        localName: 'ItemHasKnownEntity',
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'entityType',
            required: true,
            attributeName: {
              localPart: 'EntityType'
            },
            type: 'attribute'
          }, {
            name: 'regExFilter',
            attributeName: {
              localPart: 'RegExFilter'
            },
            type: 'attribute'
          }, {
            name: 'filterName',
            attributeName: {
              localPart: 'FilterName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ExecuteFunction',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ExecuteFunction'
        },
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'functionName',
            required: true,
            elementName: {
              localPart: 'FunctionName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }]
      }, {
        localName: 'CommandSurface',
        propertyInfos: [{
            name: 'customTab',
            required: true,
            collection: true,
            elementName: 'CustomTab',
            typeInfo: '.CustomTab'
          }]
      }, {
        localName: 'LocaleTokenOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'LocaleTokenOverride'
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
        localName: 'MailHost',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'MailHost'
        },
        baseTypeInfo: '.Host',
        propertyInfos: [{
            name: 'desktopFormFactor',
            elementName: {
              localPart: 'DesktopFormFactor',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.FormFactor'
          }]
      }, {
        localName: 'Event',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Event'
        },
        propertyInfos: [{
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'Type'
            },
            type: 'attribute'
          }, {
            name: 'functionExecution',
            required: true,
            typeInfo: '.EventFunctionExecutionType',
            attributeName: {
              localPart: 'FunctionExecution'
            },
            type: 'attribute'
          }, {
            name: 'functionName',
            required: true,
            attributeName: {
              localPart: 'FunctionName'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'UIControl',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'UIControl'
        },
        propertyInfos: [{
            name: 'overriddenByRibbonApi',
            elementName: {
              localPart: 'OverriddenByRibbonApi',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: 'Boolean'
          }, {
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'supertip',
            required: true,
            elementName: {
              localPart: 'Supertip',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
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
        localName: 'ConnectedServiceControlsScopes',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ConnectedServiceControlsScopes'
        },
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Scope',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }]
      }, {
        localName: 'Runtime',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Runtime'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
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
        localName: 'MenuItems',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'MenuItems'
        },
        propertyInfos: [{
            name: 'item',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Item',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.MenuItem'
          }]
      }, {
        localName: 'Button',
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: 'Action',
            typeInfo: '.Action'
          }]
      }, {
        localName: 'MessageComposeCommandSurface',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'MessageComposeCommandSurface'
        },
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
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
        localName: 'Rule',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Rule'
        }
      }, {
        localName: 'RuntimeOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'RuntimeOverride'
        },
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
        localName: 'RequirementsTokenOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'RequirementsTokenOverride'
        },
        propertyInfos: [{
            name: 'requirements',
            required: true,
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Requirements'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'Value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TaskPaneApp',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'TaskPaneApp'
        },
        baseTypeInfo: '.OfficeApp',
        propertyInfos: [{
            name: 'requirements',
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Requirements'
          }, {
            name: 'defaultSettings',
            required: true,
            elementName: {
              localPart: 'DefaultSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.TaskPaneAppSettings'
          }, {
            name: 'permissions',
            required: true,
            elementName: {
              localPart: 'Permissions',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            values: ['Restricted', 'ReadDocument', 'ReadAllDocument', 'WriteDocument', 'ReadWriteDocument']
          }, {
            name: 'dictionary',
            elementName: {
              localPart: 'Dictionary',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Dictionary'
          }, {
            name: 'versionOverrides',
            elementName: {
              localPart: 'VersionOverrides',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.VersionOverridesV10'
          }, {
            name: 'extendedOverrides',
            elementName: {
              localPart: 'ExtendedOverrides',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ExtendedOverrides'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
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
        localName: 'OfficeTab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'OfficeTab'
        },
        baseTypeInfo: '.Tab'
      }, {
        localName: 'LongResourceReference',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/officeappbasictypes\/1.0',
          localPart: 'LongResourceReference'
        },
        baseTypeInfo: '.ResourceReference'
      }, {
        localName: 'MessageReadCommandSurface',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'MessageReadCommandSurface'
        },
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
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
        localName: 'ContentAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ContentAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'requestedWidth',
            elementName: {
              localPart: 'RequestedWidth',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Int'
          }, {
            name: 'requestedHeight',
            elementName: {
              localPart: 'RequestedHeight',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Int'
          }]
      }, {
        localName: 'Menu',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Menu'
        },
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'items',
            required: true,
            elementName: {
              localPart: 'Items',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.MenuItems'
          }]
      }, {
        localName: 'WebAppAuthorizations',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'WebAppAuthorizations'
        },
        propertyInfos: [{
            name: 'authorization',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Authorization',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.WebAppAuthorization'
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
        localName: 'Tab',
        propertyInfos: [{
            name: 'group',
            required: true,
            collection: true,
            elementName: 'Group',
            typeInfo: '.Group'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'OfficeMenu',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'OfficeMenu'
        },
        propertyInfos: [{
            name: 'control',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Control',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
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
        localName: 'VersionOverridesV10',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'VersionOverridesV1_0'
        },
        propertyInfos: [{
            name: 'description',
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.LongResourceReference'
          }, {
            name: 'requirements',
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Requirements'
          }, {
            name: 'hosts',
            elementName: {
              localPart: 'Hosts',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Hosts'
          }, {
            name: 'resources',
            elementName: {
              localPart: 'Resources',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Resources'
          }, {
            name: 'webApplicationInfo',
            elementName: {
              localPart: 'WebApplicationInfo',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.WebApplicationInfo'
          }, {
            name: 'equivalentAddins',
            elementName: {
              localPart: 'EquivalentAddins',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.EquivalentAddins'
          }, {
            name: 'connectedServiceControls',
            elementName: {
              localPart: 'ConnectedServiceControls',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ConnectedServiceControls'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'ItemEdit',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemEdit'
        },
        baseTypeInfo: '.FormType',
        propertyInfos: [{
            name: 'desktopSettings',
            required: true,
            elementName: {
              localPart: 'DesktopSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemEditMailAppSettings'
          }, {
            name: 'tabletSettings',
            elementName: {
              localPart: 'TabletSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemEditMailAppSettings'
          }, {
            name: 'phoneSettings',
            elementName: {
              localPart: 'PhoneSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemEditMailAppSettings'
          }]
      }, {
        localName: 'ExecuteFunction',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ExecuteFunction'
        },
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'functionName',
            required: true,
            elementName: {
              localPart: 'FunctionName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            }
          }]
      }, {
        localName: 'CustomTab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'CustomTab'
        },
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'MessageComposeCommandSurface',
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'ItemRead',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemRead'
        },
        baseTypeInfo: '.FormType',
        propertyInfos: [{
            name: 'desktopSettings',
            required: true,
            elementName: {
              localPart: 'DesktopSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemReadDesktopMailAppSettings'
          }, {
            name: 'tabletSettings',
            elementName: {
              localPart: 'TabletSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemReadTabletMailAppSettings'
          }, {
            name: 'phoneSettings',
            elementName: {
              localPart: 'PhoneSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ItemReadPhoneMailAppSettings'
          }]
      }, {
        localName: 'WebApplicationScopes',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'WebApplicationScopes'
        },
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Scope',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
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
        localName: 'ItemHasAttachment',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemHasAttachment'
        },
        baseTypeInfo: '.Rule'
      }, {
        localName: 'WebApplicationInfo',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'WebApplicationInfo'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            elementName: {
              localPart: 'Id',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'msaId',
            elementName: {
              localPart: 'MsaId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'resource',
            required: true,
            elementName: {
              localPart: 'Resource',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'scopes',
            required: true,
            elementName: {
              localPart: 'Scopes',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.WebApplicationScopes'
          }, {
            name: 'authorizations',
            elementName: {
              localPart: 'Authorizations',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.WebAppAuthorizations'
          }]
      }, {
        localName: 'OfficeGroup',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'OfficeGroup'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Page',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Page'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'MenuItem',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'MenuItem'
        },
        baseTypeInfo: '.UIControlWithOptionalIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: {
              localPart: 'Action',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Action'
          }, {
            name: 'enabled',
            elementName: {
              localPart: 'Enabled',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: 'Boolean'
          }]
      }, {
        localName: 'FormType',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'FormType'
        }
      }, {
        localName: 'MobileAppointmentOrganizerCommandSurface',
        baseTypeInfo: '.MobileCommandSurfaceExtensionPoint'
      }, {
        localName: 'Button',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Button'
        },
        baseTypeInfo: '.UIControlWithIcon',
        propertyInfos: [{
            name: 'action',
            required: true,
            elementName: {
              localPart: 'Action',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.Action'
          }, {
            name: 'enabled',
            elementName: {
              localPart: 'Enabled',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: 'Boolean'
          }]
      }, {
        localName: 'OfficeApp',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'OfficeApp'
        },
        propertyInfos: [{
            name: 'id',
            required: true,
            elementName: {
              localPart: 'Id',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }, {
            name: 'alternateId',
            elementName: {
              localPart: 'AlternateId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }, {
            name: 'version',
            required: true,
            elementName: {
              localPart: 'Version',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }, {
            name: 'providerName',
            required: true,
            elementName: {
              localPart: 'ProviderName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }, {
            name: 'defaultLocale',
            required: true,
            elementName: {
              localPart: 'DefaultLocale',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            }
          }, {
            name: 'displayName',
            required: true,
            elementName: {
              localPart: 'DisplayName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ShortLocaleAwareSetting'
          }, {
            name: 'description',
            required: true,
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.LongLocaleAwareSetting'
          }, {
            name: 'iconUrl',
            elementName: {
              localPart: 'IconUrl',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'highResolutionIconUrl',
            elementName: {
              localPart: 'HighResolutionIconUrl',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'supportUrl',
            elementName: {
              localPart: 'SupportUrl',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'appDomains',
            elementName: {
              localPart: 'AppDomains',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.AppDomains'
          }, {
            name: 'hosts',
            elementName: {
              localPart: 'Hosts',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Hosts'
          }]
      }, {
        localName: 'MailAppRequirements',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'MailAppRequirements'
        },
        propertyInfos: [{
            name: 'sets',
            required: true,
            elementName: {
              localPart: 'Sets',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Sets'
          }]
      }, {
        localName: 'Host'
      }, {
        localName: 'WebAppAuthorizations',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'WebAppAuthorizations'
        },
        propertyInfos: [{
            name: 'authorization',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Authorization',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.WebAppAuthorization'
          }]
      }, {
        localName: 'Module',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'commandSurface',
            required: true,
            elementName: 'CommandSurface',
            typeInfo: '.CommandSurface'
          }]
      }, {
        localName: 'Tab',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Tab'
        },
        propertyInfos: [{
            name: 'group',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Group',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Group'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ExtendedOverrides',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ExtendedOverrides'
        },
        propertyInfos: [{
            name: 'tokens',
            elementName: {
              localPart: 'Tokens',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Tokens'
          }, {
            name: 'url',
            required: true,
            attributeName: {
              localPart: 'Url'
            },
            type: 'attribute'
          }, {
            name: 'resourcesUrl',
            attributeName: {
              localPart: 'ResourcesUrl'
            },
            type: 'attribute'
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
        localName: 'CustomPane',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'CustomPane'
        },
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'requestedHeight',
            elementName: {
              localPart: 'RequestedHeight',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: 'Int'
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }, {
            name: 'rule',
            required: true,
            elementName: {
              localPart: 'Rule',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Rule'
          }, {
            name: 'disableEntityHighlighting',
            elementName: {
              localPart: 'DisableEntityHighlighting',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: 'Boolean'
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
        localName: 'Hosts',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Hosts'
        },
        propertyInfos: [{
            name: 'host',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Host',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.Host'
          }]
      }, {
        localName: 'ItemHasRegularExpressionMatch',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemHasRegularExpressionMatch'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'regExName',
            required: true,
            attributeName: {
              localPart: 'RegExName'
            },
            type: 'attribute'
          }, {
            name: 'regExValue',
            required: true,
            attributeName: {
              localPart: 'RegExValue'
            },
            type: 'attribute'
          }, {
            name: 'propertyName',
            required: true,
            typeInfo: '.PropertyName',
            attributeName: {
              localPart: 'PropertyName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CustomPane',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'requestedHeight',
            elementName: 'RequestedHeight',
            typeInfo: 'Int'
          }, {
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }, {
            name: 'rule',
            required: true,
            elementName: 'Rule',
            typeInfo: '.Rule'
          }, {
            name: 'disableEntityHighlighting',
            elementName: 'DisableEntityHighlighting',
            typeInfo: 'Boolean'
          }]
      }, {
        localName: 'UIControlWithIcon',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'UIControlWithIcon'
        },
        baseTypeInfo: '.UIControl',
        propertyInfos: [{
            name: 'icon',
            required: true,
            elementName: {
              localPart: 'Icon',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.IconList'
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
        localName: 'CustomTab',
        baseTypeInfo: '.Tab',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }]
      }, {
        localName: 'Rule',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Rule'
        }
      }, {
        localName: 'ItemReadPhoneMailAppSettings',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ItemReadPhoneMailAppSettings'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }]
      }, {
        localName: 'Methods',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Methods'
        },
        propertyInfos: [{
            name: 'method',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Method',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Requirement'
          }]
      }, {
        localName: 'LongLocaleAwareSetting',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'LongLocaleAwareSetting'
        },
        propertyInfos: [{
            name: 'override',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'Override',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.LongLocaleOverride'
          }, {
            name: 'defaultValue',
            required: true,
            attributeName: {
              localPart: 'DefaultValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'MenuItems',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'MenuItems'
        },
        propertyInfos: [{
            name: 'item',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Item',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.MenuItem'
          }]
      }, {
        localName: 'ContentApp',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'ContentApp'
        },
        baseTypeInfo: '.OfficeApp',
        propertyInfos: [{
            name: 'requirements',
            elementName: {
              localPart: 'Requirements',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.Requirements'
          }, {
            name: 'defaultSettings',
            required: true,
            elementName: {
              localPart: 'DefaultSettings',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ContentAppSettings'
          }, {
            name: 'permissions',
            required: true,
            elementName: {
              localPart: 'Permissions',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            values: ['Restricted', 'ReadDocument', 'WriteDocument', 'ReadWriteDocument']
          }, {
            name: 'allowSnapshot',
            elementName: {
              localPart: 'AllowSnapshot',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: 'Boolean'
          }, {
            name: 'versionOverrides',
            elementName: {
              localPart: 'VersionOverrides',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            },
            typeInfo: '.VersionOverridesV10'
          }, {
            name: 'any',
            mixed: false,
            type: 'anyElement'
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
        localName: 'EquivalentAddin',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'EquivalentAddin'
        },
        propertyInfos: [{
            name: 'progId',
            elementName: {
              localPart: 'ProgId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'displayName',
            elementName: {
              localPart: 'DisplayName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'fileName',
            elementName: {
              localPart: 'FileName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }, {
            name: 'type',
            required: true,
            elementName: {
              localPart: 'Type',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            }
          }]
      }, {
        localName: 'MobileGroup',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'control',
            required: true,
            collection: true,
            elementName: 'Control',
            typeInfo: '.MobileButton'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
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
        localName: 'AppointmentAttendeeCommandSurface',
        baseTypeInfo: '.CommandSurfaceExtensionPoint'
      }, {
        localName: 'ItemHasAttachment',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ItemHasAttachment'
        },
        baseTypeInfo: '.Rule'
      }, {
        localName: 'Metadata',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Metadata'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'AllFormFactors',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'AllFormFactors'
        },
        propertyInfos: [{
            name: 'extensionPoint',
            elementName: {
              localPart: 'ExtensionPoint',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.CustomFunctions'
          }]
      }, {
        localName: 'ItemHasRegularExpressionMatch',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ItemHasRegularExpressionMatch'
        },
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'regExName',
            required: true,
            attributeName: {
              localPart: 'RegExName'
            },
            type: 'attribute'
          }, {
            name: 'regExValue',
            required: true,
            attributeName: {
              localPart: 'RegExValue'
            },
            type: 'attribute'
          }, {
            name: 'propertyName',
            required: true,
            typeInfo: '.PropertyName',
            attributeName: {
              localPart: 'PropertyName'
            },
            type: 'attribute'
          }, {
            name: 'ignoreCase',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IgnoreCase'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'VersionedRequirement',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
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
        localName: 'Host',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Host'
        },
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'Name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Supertip',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Supertip'
        },
        propertyInfos: [{
            name: 'title',
            required: true,
            elementName: {
              localPart: 'Title',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'description',
            required: true,
            elementName: {
              localPart: 'Description',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.LongResourceReference'
          }]
      }, {
        localName: 'Events',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'event',
            required: true,
            elementName: 'Event',
            typeInfo: '.Event'
          }]
      }, {
        localName: 'Dictionary',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Dictionary'
        },
        propertyInfos: [{
            name: 'targetDialects',
            required: true,
            elementName: {
              localPart: 'TargetDialects',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.TargetDialects'
          }, {
            name: 'queryUri',
            required: true,
            elementName: {
              localPart: 'QueryUri',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
          }, {
            name: 'citationText',
            required: true,
            elementName: {
              localPart: 'CitationText',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ShortLocaleAwareSetting'
          }, {
            name: 'dictionaryName',
            required: true,
            elementName: {
              localPart: 'DictionaryName',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.ShortLocaleAwareSetting'
          }, {
            name: 'dictionaryHomePage',
            required: true,
            elementName: {
              localPart: 'DictionaryHomePage',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
            },
            typeInfo: '.URLLocaleAwareSetting'
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
        localName: 'Hosts',
        propertyInfos: [{
            name: 'host',
            required: true,
            collection: true,
            elementName: 'Host',
            typeInfo: '.Host'
          }]
      }, {
        localName: 'Group',
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: 'Label',
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'tooltip',
            elementName: 'Tooltip',
            typeInfo: '.ShortResourceReference'
          }, {
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
        localName: 'MessageComposeAutoRun',
        baseTypeInfo: '.ExtensionPoint',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: 'SourceLocation',
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'OfficeTab',
        baseTypeInfo: '.Tab'
      }, {
        localName: 'Rule',
        propertyInfos: [{
            name: 'highlight',
            attributeName: {
              localPart: 'Highlight'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ConnectedServiceControlsScopes',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides',
          localPart: 'ConnectedServiceControlsScopes'
        },
        propertyInfos: [{
            name: 'scope',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Scope',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
            }
          }]
      }, {
        localName: 'Token',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
          localPart: 'Token'
        },
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'Name'
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
        localName: 'ItemIs',
        baseTypeInfo: '.Rule',
        propertyInfos: [{
            name: 'itemType',
            required: true,
            attributeName: {
              localPart: 'ItemType'
            },
            type: 'attribute'
          }, {
            name: 'itemClass',
            attributeName: {
              localPart: 'ItemClass'
            },
            type: 'attribute'
          }, {
            name: 'includeSubClasses',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'IncludeSubClasses'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ExtensionPoint',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'ExtensionPoint'
        }
      }, {
        localName: 'Script',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides',
          localPart: 'Script'
        },
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'URLLocaleOverride',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1',
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
        localName: 'ShowTaskpane',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ShowTaskpane'
        },
        baseTypeInfo: '.Action',
        propertyInfos: [{
            name: 'sourceLocation',
            required: true,
            elementName: {
              localPart: 'SourceLocation',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.URLResourceReference'
          }]
      }, {
        localName: 'ExtensionPoint',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'ExtensionPoint'
        }
      }, {
        localName: 'Group',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides',
          localPart: 'Group'
        },
        propertyInfos: [{
            name: 'label',
            required: true,
            elementName: {
              localPart: 'Label',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'tooltip',
            elementName: {
              localPart: 'Tooltip',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
            typeInfo: '.ShortResourceReference'
          }, {
            name: 'control',
            required: true,
            collection: true,
            elementName: {
              localPart: 'Control',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
            },
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
        type: 'enumInfo',
        localName: 'ItemFormType',
        values: ['Read', 'Edit', 'ReadOrEdit']
      }, {
        type: 'enumInfo',
        localName: 'ItemType',
        values: ['Message', 'Appointment']
      }, {
        type: 'enumInfo',
        localName: 'KnownEntityType',
        values: ['MeetingSuggestion', 'TaskSuggestion', 'Address', 'Url', 'PhoneNumber', 'EmailAddress', 'Contact']
      }, {
        type: 'enumInfo',
        localName: 'PropertyName',
        values: ['Subject', 'BodyAsPlaintext', 'BodyAsHTML', 'SenderSMTPAddress']
      }, {
        type: 'enumInfo',
        localName: 'LaunchEventSendMode',
        values: ['PromptUser', 'SoftBlock', 'Block']
      }, {
        type: 'enumInfo',
        localName: 'EventFunctionExecutionType',
        values: ['synchronous', 'asynchronous']
      }, {
        type: 'enumInfo',
        localName: 'PropertyName',
        values: ['Subject', 'BodyAsPlaintext', 'BodyAsHTML', 'SenderSMTPAddress']
      }, {
        type: 'enumInfo',
        localName: 'LogicalOperator',
        values: ['And', 'Or']
      }, {
        type: 'enumInfo',
        localName: 'PropertyName',
        values: ['Subject', 'BodyAsPlaintext', 'BodyAsHTML', 'SenderSMTPAddress']
      }, {
        type: 'enumInfo',
        localName: 'EventFunctionExecutionType',
        values: ['synchronous', 'asynchronous', 'synchronousUserDismissable']
      }, {
        type: 'enumInfo',
        localName: 'LogicalOperator',
        values: ['And', 'Or']
      }, {
        type: 'enumInfo',
        localName: 'LogicalOperator',
        values: ['And', 'Or']
      }],
    elementInfos: [{
        typeInfo: '.VersionOverridesV11',
        elementName: 'VersionOverrides'
      }, {
        typeInfo: '.VersionOverridesV10',
        elementName: {
          localPart: 'VersionOverrides',
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/mailappversionoverrides'
        }
      }, {
        typeInfo: '.OfficeApp',
        elementName: {
          localPart: 'OfficeApp',
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/appforoffice\/1.1'
        }
      }, {
        typeInfo: '.VersionOverridesV10',
        elementName: {
          localPart: 'VersionOverrides',
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/contentappversionoverrides'
        }
      }, {
        typeInfo: '.VersionOverridesV10',
        elementName: {
          localPart: 'VersionOverrides',
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/taskpaneappversionoverrides'
        }
      }]
  };
  return {
    OfficeAppManifest: OfficeAppManifest
  };
};
if (typeof define === 'function' && define.amd) {
  define([], OfficeAppManifest_Module_Factory);
}
else {
  var OfficeAppManifest_Module = OfficeAppManifest_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.OfficeAppManifest = OfficeAppManifest_Module.OfficeAppManifest;
  }
  else {
    var OfficeAppManifest = OfficeAppManifest_Module.OfficeAppManifest;
  }
}