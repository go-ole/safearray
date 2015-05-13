package safearray

import "github.com/go-ole/com"

func NewVector(variantType com.VariantType, slice []interface{}) (safearray *COMArray, err error) {
	safearray, err = CreateArrayVector(variantType, 0, uint32(len(slice)))
	if err != nil {
		return
	}

	return UnmarshalArray(safearray, slice)
}

func NewVectorWithFlags(variantType com.VariantType, flags com.SafeArrayMask, slice []interface{}) (safearray *COMArray, err error) {
	safearray, err = CreateArrayVectorEx(variantType, 0, uint32(len(slice)), uintptr(flags))
	if err != nil {
		return
	}

	return UnmarshalArray(safearray, slice)
}

// FromByteArray creates SafeArray from byte array.
func FromByteArray(slice []byte) (safearray *COMArray, err error) {
	return NewVector(com.UInteger8VariantType, slice)
}
