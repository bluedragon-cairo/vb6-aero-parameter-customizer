API usage:
 private static extern void DwmSetColorizationParameters(ref DWM_COLORIZATION_PARAMS parameters, bool temporary);

 DWM_COLORIZATION_PARAMS parameters:
  A typedef of colorization parameters.
 bool temporary:
  If true or 1, the set parameters aren't written to registry and will be discarded apon DWM restart
