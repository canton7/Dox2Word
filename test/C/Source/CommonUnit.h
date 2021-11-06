/**
 * @defgroup CommonUnit Common - Unit
 * @ingroup Common
 * 
 * Common Unit is a Unit in the Common module.
 * 
 * <strong>Bold and <em>italic and @c typewriter </em></strong>
 * 
 * @warning This is a very important warning!
 *
 * A list of things:
 *  1. Thing 1
 *    1. Thing 1.1
 *  2. Thing 2
 * 
 * @dot
 * digraph example {
 *     node [shape=record, fontname=Helvetica, fontsize=10];
 *     b [ label="class B" URL="\ref B"];
 *     c [ label="class C" URL="\ref C"];
 *     b -> c [ arrowhead="open", style="dashed" ];
 * }
 * @enddot
 * 
 * @{
 * @file
 */

#include <stdint.h>
#include <stdbool.h>

/// This is a macro
///
/// @param x Thing
/// @returns @c x doubled
#define COMMON_UNIT_MACRO(x) (uint32_t)(x * 2)

/// This is a constant
#define COMMON_UNIT_CONST (uint32_t)3

/**
 * This is a global variable
 */
extern bool CommonUnit_Global;

/**
 * A test structure
 * 
 * More description of the test structure
 */
typedef struct CommonUnit_Struct_tag
{
    /// Member 'i'
    ///
    /// This member is so important it gets a detailed description.
    ///
    /// With two paragraphs in it.
	int32_t i;

    bool j; ///< Just 'j'
} CommonUnit_Struct_t;

/**
 * A test enum
 */
typedef enum CommonUnit_Enum_tag
{
    /// This one doesn't have a value
    CommonUnit_Enum_One,
    CommonUnit_Enum_Two = 2,    ///< This one has a value, and inline does

    /// This one has a brief description.
    /// And then a more detailed description.
    CommonUnit_Enum_Three,

    /// This one's dull
    CommonUnit_Enum_Four = CommonUnit_Enum_Three
} CommonUnit_Enum_t;

/**
 * A test union
 */
typedef union CommonUnit_Union_tag
{
    /// Union member 1
    uint8_t one;

    /// Union member 2
    CommonUnit_Enum_t two;
} CommonUnit_Union_t;

/**
 * A struct containing an anonymous union
 */
typedef struct CommonUnit_StructWithUnion_tag
{
    /// The type discriminator
    uint8_t type;
    union
    {
        uint16_t one : 4; ///< A bitfield
        uint32_t two;   ///< A normal field
    };
} CommonUnit_StructWithUnion_t;

/**
 * This is a custom typedef
 */
typedef uint8_t CommonUnit_TypedefType_t;

/**
 * This is the common unit's test file
 * 
 * This is a longer description of the function.
 * 
 * And another paragraph of description.
 * 
 * @param p Some parameters
 * @returns A value
 * 
 * Some more text afterwards.
 */
void CommonUnit_Test(const uint8_t p);

/**
 * Another function in CommonUnit
 * 
 * Usage:
 * @code{.c}
 * int result = CommonUnit_Another();
 * Foo x = rand();
 * @endcode
 *
 * @param in In documentation @c yay
 * @param foo Foo documentation
 * @returns A value
 * @retval 1 Some return value
 * @retval 2 Another return value
 */
CommonUnit_TypedefType_t CommonUnit_Another(const CommonUnit_Struct_t* in, uint8_t* const foo);

/**
 * Thingy docs
 */
void Thingy(void);

/// @}
