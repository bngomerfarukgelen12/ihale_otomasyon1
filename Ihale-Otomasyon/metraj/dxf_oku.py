import ezdxf

doc = ezdxf.readfile("dosya.dxf")
msp = doc.modelspace()

layer_data = {}

for e in msp:
    layer = e.dxf.layer

    if layer not in layer_data:
        layer_data[layer] = 0

    # LINE
    if e.dxftype() == "LINE":
        layer_data[layer] += e.length()

    # LWPOLYLINE
    elif e.dxftype() == "LWPOLYLINE":
        layer_data[layer] += e.length()

for layer, length in layer_data.items():
    print(layer, length)