SITE_URL=$1
EMAIL=$2
PASS=$3

npm i -g @pnp/office365-cli
o365 version

o365 spo login $SITE_URL --authType password --userName $EMAIL --password $PASS

o365 spo app add --filePath "./nakkeerann.globalspfxsoln/SPFx sppkg/globalspfxsoln.sppkg" --overwrite --verbose