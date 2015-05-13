package safearray

import (
	"unsafe"

	"github.com/go-ole/com"
)

func NewVector(variantType com.VariantType, slice []interface{}) (safearray *COMArray, err error) {
	safearray, err = CreateArrayVector(variantType, 0, uint32(len(slice)))
	if err != nil {
		return
	}

	return UnmarshalArray(safearray, &slice)
}

func NewVectorWithFlags(variantType com.VariantType, flags com.SafeArrayMask, slice []interface{}) (safearray *COMArray, err error) {
	safearray, err = CreateArrayVectorEx(variantType, 0, uint32(len(slice)), uintptr(flags))
	if err != nil {
		return
	}

	return UnmarshalArray(safearray, &slice)
}

func UnmarshalArray(safearray *COMArray, slice []interface{}) (safearray *COMArray, err error) {
	if safearray == nil {
		err = SafeArrayVectorFailedError
		return
	}

	for pos, val := range slice {
		PutElement(safearray, int64(pos), uintptr(unsafe.Pointer(&val)))
	}

	return
}
