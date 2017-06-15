{$IFDEF WIN64}
  {$IFDEF DEBUG}
    {$LIBSUFFIX '_64_debug'}
  {$ELSE}
    {$LIBSUFFIX '_64'}
  {$ENDIF}
{$ELSE}
  {$IFDEF DEBUG}
    {$LIBSUFFIX '_32_debug'}
  {$ELSE}
    {$LIBSUFFIX '_32'}
  {$ENDIF}
{$ENDIF}
