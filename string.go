package safearray

import (
	"unsafe"

	"github.com/go-ole/com"
)

// StringArray is an helper for converting SafeArray to strings.
//
// SafeArray can be convert to StringArray and ToArray() can then be called.
type StringArray *Array

// ToArray converts SafeArray to Go string array.
func (ssa *StringArray) ToArray() []string {
	return ToStringArray(*Array(ssa))
}

// GetElementString is helper function for retrieving string from element.
func GetElementString(safearray *COMArray, index int64) (str string, err error) {
	var element *int16
	err = PutElementIn(safearray, index, &element)
	str = com.BstrToString(*(**uint16)(unsafe.Pointer(&element)))
	com.SysFreeString(element)
	return
}

// ToStringArray converts SafeArray object to string array.
//
// BUG: Only gets the first dimension.
func ToStringArray(safearray *Array) (strings []string) {
	totalElements, _ := safearray.TotalElements(0)
	strings = make([]string, totalElements)

	for i := int64(0); i < totalElements; i++ {
		strings[int32(i)], _ = GetElementString(safearray.Array, i)
	}

	return
}
