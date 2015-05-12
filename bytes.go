package safearray

import (
	"unsafe"

	"github.com/go-ole/com"
)

// ByteArray is an helper for converting SafeArray to a byte array.
//
// SafeArray can be convert to ByteArray and ToArray() can then be called.
type ByteArray *Array

// ToArray converts SafeArray to Go string array.
func (bsa *ByteArray) ToArray() []byte {
	return ToByteArray(*Array(bsa))
}

// FromByteSlice creates SafeArray from byte array.
func FromByteSlice(slice []byte) *COMArray {
	array, _ := CreateArrayVector(com.UInteger8VariantType, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []byte to SafeArray")
	}

	for i, v := range slice {
		PutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

// ToByteArray converts SafeArray to byte array.
//
// BUG: Only gets the first dimension.
func ToByteArray(safearray *Array) (bytes []byte) {
	totalElements, _ := safearray.TotalElements(0)
	bytes = make([]byte, totalElements)

	for i := int64(0); i < totalElements; i++ {
		var b byte
		_ := PutElementIn(safearray.Array, i, &b)
		bytes[int32(i)] = b
	}

	return
}
