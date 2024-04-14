from django.db import models

# Create your models here.
class CV(models.Model):
    file = models.FileField(upload_to='cv_files/')
