/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Maps class properties to Excel fields.
    /// </summary>
    /// <typeparam name="T">The <see cref="Type"/> of class to map.</typeparam>
    public abstract class ExcelClassMap<T> : ExcelClassMap
    {
        /// <summary>
        /// Constructs the row object using the given expression.
        /// </summary>
        /// <param name="expression">The expression.</param>
        protected virtual void ConstructUsing(
            Expression<Func<T>> expression)
        {
            Constructor = ReflectionHelper.GetConstructor(expression);
        }

        /// <summary>
        /// Maps a property to a Excel field.
        /// </summary>
        /// <param name="property">Property to map</param>
        /// <returns>The property mapping.</returns>
        private ExcelPropertyMap Map(
            PropertyInfo property)
        {
            var existingMap = PropertyMaps.SingleOrDefault(m => m.Data.Property == property);
            if (existingMap != null) {
                return existingMap;
            }

            var propertyMap = new ExcelPropertyMap(property);
            propertyMap.Data.Index = GetMaxIndex() + 1;
            PropertyMaps.Add(propertyMap);

            return propertyMap;
        }

        /// <summary>
        /// Maps a property to a Excel field.
        /// </summary>
        /// <param name="expression">The property to map.</param>
        /// <returns>The property mapping.</returns>
        protected ExcelPropertyMap Map(
            Expression<Func<T, object>> expression)
        {
            return Map(ReflectionHelper.GetProperty(expression));
        }

        /// <summary>
        /// Maps a property to a Excel field by name
        /// </summary>
        /// <param name="name">Name of the property to map</param>
        /// <returns>The property mapping.</returns>
        protected ExcelPropertyMap Map(
            string name)
        {
            return Map(typeof(T).GetProperty(name));
        }

        /// <summary>
        /// Determines if a column that is mapped is actually imported in the Excel file
        /// </summary>
        /// <param name="expression">The property to map.</param>
        /// <param name="importedColumns">List of mapped columns to check against</param>
        /// <returns>The property mapping.</returns>
        public static bool IsImported(
            Expression<Func<T, object>> expression,
            List<PropertyInfo> importedColumns)
        {
            var property = ReflectionHelper.GetProperty(expression);
            return importedColumns.FirstOrDefault(p => p == property) != null;
        }

        /// <summary>
        /// Maps a property to another class map.
        /// </summary>
        /// <typeparam name="TClassMap">The type of the class map.</typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns>The reference mapping for the property.</returns>
        protected ExcelPropertyReferenceMap References<TClassMap>(
            Expression<Func<T, object>> expression)
            where TClassMap : ExcelClassMap
        {
            return References(typeof(TClassMap), expression);
        }

        /// <summary>
        /// Maps a property to another class map.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="expression">The expression.</param>
        /// <returns>The reference mapping for the property</returns>
        protected ExcelPropertyReferenceMap References(
            Type type,
            Expression<Func<T, object>> expression)
        {
            var property = ReflectionHelper.GetProperty(expression);
            var map = (ExcelClassMap)ReflectionHelper.CreateInstance(type);
            map.ReIndex(GetMaxIndex() + 1);
            var reference = new ExcelPropertyReferenceMap(property, map);
            ReferenceMaps.Add(reference);
            return reference;
        }
    }
}
