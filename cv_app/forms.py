from django import forms

class MultipleFileInput(forms.FileInput):
    # Override the render method to include 'multiple' attribute
    def render(self, name, value, attrs=None, renderer=None):
        attrs['multiple'] = 'multiple'
        return super().render(name, value, attrs=attrs, renderer=renderer)

class FileUploadForm(forms.Form):
    files = forms.FileField(widget=MultipleFileInput)
