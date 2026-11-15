type: minor

La propagación de la flota regenera el documento operativo del consumer con el canon que propaga y lo commitea en el mismo commit (#1152). Un consumer que declara la fase 5 en su contrato de adopción deja de recibir un canon nuevo con el documento viejo: el documento se genera con `ci-pattern adoption generate-doc` a partir de los datos reales del consumer, viaja en el mismo trozo de la cadena que el cambio del canon y su `sha256` se recalcula en el manifiesto de gobierno cuando el consumer lo sella (HR-46, HR-72).
