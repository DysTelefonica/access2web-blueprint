type: minor

HR-44 cambia de modelo (#665 fase B): cada edición material de una skill añade un fragmento válido en changelog.d/ y ya no sube metadata.version en el PR. El gate version-bump se retira; el gate changelog-fragment se cablea en CI. La versión la sube el PR de release.
