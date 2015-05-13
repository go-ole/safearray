// +build !windows

package safearray

import "unsafe"
import "github.com/go-ole/com"

// AccessData returns raw array.
//
// AKA: SafeArrayAccessData in Windows API.
func AccessData(safearray *COMArray) (uintptr, error) {
	return uintptr(unsafe.Pointer(safearray)), NotImplementedError
}

// UnaccessData releases raw array.
//
// AKA: SafeArrayUnaccessData in Windows API.
func UnaccessData(safearray *COMArray) error {
	return NotImplementedError
}

// AllocateArrayData allocates SafeArray.
//
// AKA: SafeArrayAllocData in Windows API.
func AllocateArrayData(safearray *COMArray) error {
	return NotImplementedError
}

// AllocateArrayDescriptor allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptor in Windows API.
func AllocateArrayDescriptor(dimensions uint32) (*COMArray, error) {
	return nil, NotImplementedError
}

// AllocateArrayDescriptorEx allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptorEx in Windows API.
func AllocateArrayDescriptorEx(variantType com.VariantType, dimensions uint32) (*COMArray, error) {
	return nil, NotImplementedError
}

// Duplicate returns copy of SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func Duplicate(original *COMArray) (*COMArray, error) {
	return nil, NotImplementedError
}

// DuplicateData duplicates SafeArray into another SafeArray object.
//
// AKA: SafeArrayCopyData in Windows API.
func DuplicateData(original, duplicate *COMArray) error {
	return NotImplementedError
}

// CreateArray creates SafeArray.
//
// AKA: SafeArrayCreate in Windows API.
func CreateArray(variantType com.VariantType, dimensions uint32, bounds *Bounds) (*COMArray, error) {
	return nil, NotImplementedError
}

// CreateArrayEx creates SafeArray.
//
// AKA: SafeArrayCreateEx in Windows API.
func CreateArrayEx(variantType com.VariantType, dimensions uint32, bounds *Bounds, extra uintptr) (*COMArray, error) {
	return nil, NotImplementedError
}

// CreateArrayVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func CreateArrayVector(variantType com.VariantType, lowerBound int32, length uint32) (*COMArray, error) {
	return nil, NotImplementedError
}

// CreateArrayVectorEx creates SafeArray.
//
// AKA: SafeArrayCreateVectorEx in Windows API.
func CreateArrayVectorEx(variantType com.VariantType, lowerBound int32, length uint32, extra uintptr) (*COMArray, error) {
	return nil, NotImplementedError
}

// Destroy destroys SafeArray object.
//
// AKA: SafeArrayDestroy in Windows API.
func Destroy(safearray *COMArray) error {
	return NotImplementedError
}

// DestroyData destroys SafeArray object.
//
// AKA: SafeArrayDestroyData in Windows API.
func DestroyData(safearray *COMArray) error {
	return NotImplementedError
}

// DestroyDescriptor destroys SafeArray object.
//
// AKA: SafeArrayDestroyDescriptor in Windows API.
func DestroyDescriptor(safearray *COMArray) error {
	return NotImplementedError
}

// GetDimensions is the amount of dimensions in the SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetDim in Windows API.
func GetDimensions(safearray *COMArray) (uint32, error) {
	return *uint32(0), NotImplementedError
}

// GetElementSize is the element size in bytes.
//
// AKA: SafeArrayGetElemsize in Windows API.
func GetElementSize(safearray *COMArray) (uint32, error) {
	return *uint32(0), NotImplementedError
}

// GetElement retrieves element at given index.
func GetElement(safearray *COMArray, index int64) (interface{}, error) {
	return nil, NotImplementedError
}

// PutElementIn retrieves element value at given index.
//
// AKA: SafeArrayGetElement in Windows API.
func PutElementIn(safearray *COMArray, index int64, element interface{}) error {
	return NotImplementedError
}

// GetInterfaceID is the InterfaceID of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func GetInterfaceID(safearray *COMArray) (*com.GUID, error) {
	return nil, NotImplementedError
}

// GetLowerBound returns lower bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetLBound in Windows API.
func GetLowerBound(safearray *COMArray, dimension uint32) (int64, error) {
	return int64(0), NotImplementedError
}

// GetUpperBound returns upper bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetUBound in Windows API.
func GetUpperBound(safearray *COMArray, dimension uint32) (int64, error) {
	return int64(0), NotImplementedError
}

// GetVariantType returns data type of SafeArray.
//
// AKA: SafeArrayGetVartype in Windows API.
func GetVariantType(safearray *COMArray) (com.VariantType, error) {
	return com.NullVariantType, NotImplementedError
}

// Lock locks SafeArray for reading to modify SafeArray.
//
// This must be called during some calls to ensure that another process does not
// read or write to the SafeArray during editing.
//
// AKA: SafeArrayLock in Windows API.
func Lock(safearray *COMArray) error {
	return NotImplementedError
}

// Unlock unlocks SafeArray for reading.
//
// AKA: SafeArrayUnlock in Windows API.
func Unlock(safearray *COMArray) error {
	return NotImplementedError
}

// GetPointerOfIndex gets a pointer to an array element.
//
// AKA: SafeArrayPtrOfIndex in Windows API.
func GetPointerOfIndex(safearray *COMArray, index int64) (uintptr, error) {
	return uintptr(0), NotImplementedError
}

// SetInterfaceID sets the GUID of the interface for the specified safe
// array.
//
// AKA: SafeArraySetIID in Windows API.
func SetInterfaceID(safearray *COMArray, interfaceID *com.GUID) error {
	return NotImplementedError
}

// ResetDimensions changes the right-most (least significant) bound of the
// specified safe array.
//
// AKA: SafeArrayRedim in Windows API.
func ResetDimensions(safearray *COMArray, bounds *Bounds) error {
	return NotImplementedError
}

// PutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func PutElement(safearray *COMArray, index int64, element interface{}) error {
	return NotImplementedError
}

// GetRecordInfo accesses IRecordInfo info for custom types.
//
// AKA: SafeArrayGetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func GetRecordInfo(safearray *COMArray) (interface{}, error) {
	return nil, NotImplementedError
}

// SetRecordInfo mutates IRecordInfo info for custom types.
//
// AKA: SafeArraySetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func SetRecordInfo(safearray *COMArray, recordInfo interface{}) (err error) {
	return NotImplementedError
}
