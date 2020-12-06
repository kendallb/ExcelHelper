/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelHelper.Configuration
{
    ///<summary>
    /// Maps class properties to Excel fields.
    ///</summary>
    public abstract class ExcelClassMapBase
    {
        private readonly ExcelPropertyMapCollection _propertyMaps = new ExcelPropertyMapCollection();
        private readonly List<ExcelPropertyReferenceMap> _referenceMaps = new List<ExcelPropertyReferenceMap>();

        /// <summary>
        /// Gets the constructor expression.
        /// </summary>
        public virtual NewExpression Constructor { get; protected set; }

        /// <summary>
        /// The class property mappings.
        /// </summary>
        public virtual ExcelPropertyMapCollection PropertyMaps => _propertyMaps;

        /// <summary>
        /// The class property reference mappings.
        /// </summary>
        public virtual List<ExcelPropertyReferenceMap> ReferenceMaps => _referenceMaps;

        /// <summary>
        /// Allow only internal creation of ExcelClassMap.
        /// </summary>
        internal ExcelClassMapBase()
        {
        }

        /// <summary>
        /// Gets the property map for the given property expression.
        /// </summary>
        /// <typeparam name="T">The type of the class the property belongs to.</typeparam>
        /// <param name="propertyExpression">The property expression.</param>
        /// <returns>The ExcelPropertyMap for the given expression.</returns>
        public virtual ExcelPropertyMap PropertyMap<T>(
            Expression<Func<T, object>> propertyExpression)
        {
            var property = ReflectionHelper.GetProperty(propertyExpression);
            var propertyMap = _propertyMaps.Single(pm => pm.Data.Property == property);
            return propertyMap;
        }

        /// <summary>
        /// Get the largest index for the
        /// properties and references.
        /// </summary>
        /// <returns>The max index.</returns>
        internal int GetMaxIndex()
        {
            if (PropertyMaps.Count == 0 && ReferenceMaps.Count == 0) {
                return -1;
            }

            var indexes = new List<int>();
            if (PropertyMaps.Count > 0) {
                indexes.Add(PropertyMaps.Max(pm => pm.Data.Index));
            }
            indexes.AddRange(ReferenceMaps.Select(referenceMap => referenceMap.GetMaxIndex()));

            return indexes.Max();
        }

        /// <summary>
        /// Resets the indexes based on the given start index.
        /// </summary>
        /// <param name="indexStart">The index start.</param>
        /// <returns>The last index + 1.</returns>
        internal int ReIndex(
            int indexStart = 0)
        {
            foreach (var propertyMap in PropertyMaps) {
                propertyMap.Data.Index = indexStart;
                indexStart++;
            }
            foreach (var referenceMap in ReferenceMaps) {
                indexStart = referenceMap.Mapping.ReIndex(indexStart);
            }
            return indexStart;
        }

        /// <summary>
        /// Auto maps all properties for the given type. If a property
        /// is mapped again it will override the existing map.
        /// </summary>
        public virtual void AutoMap()
        {
            AutoMapInternal(this);
        }

        /// <summary>
        /// Auto maps the given map and checks for circular references as it goes.
        /// </summary>
        /// <param name="mapBase">The map to auto map.</param>
        internal static void AutoMapInternal(
            ExcelClassMapBase mapBase)
        {
            var type = mapBase.GetType().BaseType.GetGenericArguments()[0];
            if (typeof(IEnumerable).IsAssignableFrom(type)) {
                throw new ExcelConfigurationException("Types that inherit IEnumerable cannot be auto mapped. " +
                                                      "Did you accidentally call GetRecord or WriteRecord which " +
                                                      "acts on a single record instead of calling GetRecords or " +
                                                      "WriteRecords which acts on a list of records?");
            }

            // Process all the properties in this type
            foreach (var property in type.GetProperties()) {
                var propertyMap = new ExcelPropertyMap(property);
                propertyMap.Data.Index = mapBase.GetMaxIndex() + 1;
                mapBase.PropertyMaps.Add(propertyMap);
            }

            // Re-index all the properties when we are done
            mapBase.ReIndex();
        }
    }
}