from django.apps import AppConfig


class InformeConfig(AppConfig):
    name = 'informe'
    
    def ready(self):
        import informe.signals
