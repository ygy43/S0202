Namespace MyHelpers
    Public Class RodEndUnitViewModelModelBinder
        Inherits DefaultModelBinder

        ''' <summary>
        '''     ロッド先端情報を種類ごとにバインドできるように
        ''' </summary>
        ''' <param name="controllerContext"></param>
        ''' <param name="bindingContext"></param>
        ''' <param name="modelType"></param>
        ''' <returns></returns>
        Protected Overrides Function CreateModel(controllerContext As ControllerContext,
                                                 bindingContext As ModelBindingContext, modelType As Type) As Object
            Dim typeValue = bindingContext.ValueProvider.GetValue(bindingContext.ModelName & ".ModelType")
            Dim type = System.Type.GetType(modelType.Namespace & "." & typeValue.AttemptedValue, True)
            Dim model = Activator.CreateInstance(type)
            bindingContext.ModelMetadata = ModelMetadataProviders.Current.GetMetadataForType(Function() model, type)
            Return model
        End Function

        'Implements IModelBinder

        'Public Overloads Function BindModel(controllerContext As ControllerContext, bindingContext As ModelBindingContext) As Object Implements IModelBinder.BindModel
        '    Dim typeValue = bindingContext.ValueProvider.GetValue(bindingContext.ModelName & ".ModelType")
        '    Dim type = System.Type.GetType("S0202.ViewModels.Options." & typeValue.AttemptedValue)
        '    Dim model = Activator.CreateInstance(type)
        '    bindingContext.ModelMetadata = ModelMetadataProviders.Current.GetMetadataForType(Function() model, type)
        '    Return model
        'End Function
    End Class
End Namespace