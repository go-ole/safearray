package safearray

import "github.com/go-ole/com"

// COMArray is how COM handles arrays.
type COMArray struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uint32
	Bounds       uintptr
}

// Bounds defines the array boundaries.
type Bounds struct {
	Elements   uint32
	LowerBound int32
}

// SafeArray storage container with helpers.
//
// It is recommended that you use this type instead of the COMArray, because
// the bounds is a pointer to the SafeArrayBounds and not referenced directly.
type Array struct {
	Array *COMArray

	// bounds contains a mirror of the COMArray.Bounds in the Go type.
	bounds []Bounds
}

// Destroy SafeArray object.
func (sa *Array) Destroy() error {
	return Destroy(sa.Array)
}

// DestroyData removes safe array data.
func (sa *Array) DestroyData() error {
	return DestroyData(sa.Array)
}

// DestroyDescriptor removes safe array descriptor.
func (sa *Array) DestroyDescriptor() error {
	return DestroyDescriptor(sa.Array)
}

// Duplicate SafeArray to another SafeArray.
//
// This copies the underlying COMArray object into another Array object.
func (sa *Array) Duplicate() (*Array, error) {
	saCopy, err := Duplicate(sa.Array)
	return &Array{saCopy}, err
}

// DuplicateDataTo takes current SafeArray Data and copies to given SafeArray.
//
// This copies the underlying COMArray data into another SafeArray
// COMArray object.
func (sa *Array) DuplicateDataTo(duplicate *Array) error {
	return DuplicateData(sa.Array, duplicate.Array)
}

// Dimensions is the total number of array of arrays.
//
// For example is dimensions returns 3, then you have:
//
//     array[0][]
//     array[1][]
//     array[2][]
//
// And so on for other lengths.
func (sa *Array) Dimensions() (uint32, error) {
	return GetDimensions(sa.Array)
}

func (sa *Array) ResetDimensions(bounds []Bounds) error {
	sa.bounds = bounds
	return ResetDimensions(sa.Array, &sa.bounds[0])
}

// ElementSize is the type's size.
func (sa *Array) ElementSize() (uint32, error) {
	return GetElementSize(sa.Array)
}

// TotalElements returns total elements for given dimension.
//
// Dimensions start at 1, this will only be corrected if you enter '0'.
func (sa *Array) TotalElements(dimension uint32) (totalElements int64, err error) {
	if dimension < 1 {
		dimension = 1
	}

	// Get array bounds
	var LowerBounds int64
	var UpperBounds int64

	LowerBounds, err = GetLowerBound(sa.Array, dimension)
	if err != nil {
		return
	}

	UpperBounds, err = GetUpperBound(sa.Array, dimension)
	if err != nil {
		return
	}

	totalElements = UpperBounds - LowerBounds + 1
	return
}

// SetElementAt with element value at index.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) SetElementAt(index int64, element interface{}) error {
	return PutElement(sa.Array, index, &element)
}

// ElementAt returns element at index.
//
// Returned value will need to be converted to the type you require, because it
// is an interface{}.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) ElementAt(index int64) (interface{}, error) {
	return GetElement(sa.Array, index)
}

// ElementFor puts element value into given element.
//
// You do not need to convert element. It will be typed to the interface. This
// is an unsafe operation. Element must be passed by reference.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) ElementFor(index int64, element interface{}) error {
	return PutElementIn(sa.Array, index, &element)
}

// SetInterfaceID sets the IID for the COM array.
//
// This is only used when serving COM arrays to clients.
func (sa *Array) SetInterfaceID(interfaceID *com.GUID) error {
	return SetInterfaceID(sa.Array, &interfaceID)
}

// InterfaceID may return the IID, if the array type is a COM object.
func (sa *Array) InterfaceID() (*com.GUID, error) {
	return GetInterfaceID(sa.Array)
}

func (sa *Array) VariantType() (varType com.VariantType, err error) {
	vt, err := GetVariantType(sa.Array)
	varType = com.VariantType(vt)
	return
}

// Lock for modification.
func (sa *Array) Lock() error {
	return Lock(sa.Array)
}

// UnlockArray for reading.
func (sa *Array) Unlock() error {
	return Unlock(sa.Array)
}

// RecordInfo retrieves IRecordInfo for SafeArray.
//
// XXX: Must implement IRecordInfo interface for this to return.
func (sa *Array) RecordInfo() (interface{}, error) {
	return GetRecordInfo(sa.Array)
}

// SetRecordInfo sets IRecordInfo for SafeArray.
//
// XXX: Must implement IRecordInfo interface for this to return.
func (sa *Array) SetRecordInfo(info interface{}) error {
	return SetRecordInfo(sa.Array, info)
}

// ToArray converts SafeArray data to arbitrary type slice.
func (sa *Array) ToArray(value interface{}) (err error) {
	// TODO: Complete.
	dimensions := GetDimensions(sa.Array)
}
