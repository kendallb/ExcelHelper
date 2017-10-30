/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Creates <see cref="TypeConverterOptions"/>.
    /// </summary>
    public static class TypeConverterOptionsFactory
    {
        private static readonly Dictionary<Type, TypeConverterOptions> _typeConverterOptions = new Dictionary<Type, TypeConverterOptions>();
        private static readonly object _locker = new object();

        /// <summary>
        /// Adds the <see cref="TypeConverterOptions"/> for the given <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The type the options are for.</param>
        /// <param name="options">The options.</param>
        public static void AddOptions(
            Type type,
            TypeConverterOptions options)
        {
            if (type == null) {
                throw new ArgumentNullException(nameof(type));
            }
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }
            lock (_locker) {
                _typeConverterOptions[type] = options;
            }
        }

        /// <summary>
        /// Get the <see cref="TypeConverterOptions"/> for the given <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The type the options are for.</param>
        /// <param name="defaultCulture">Default culture to use if not specified</param>
        /// <returns>The options for the given type.</returns>
        public static TypeConverterOptions GetOptions(
            Type type,
            CultureInfo defaultCulture)
        {
            if (type == null) {
                throw new ArgumentNullException();
            }
            lock (_locker) {
                TypeConverterOptions options;
                if (!_typeConverterOptions.TryGetValue(type, out options)) {
                    options = new TypeConverterOptions();
                }
                if (options.CultureInfo == null) {
                    options.CultureInfo = defaultCulture;
                }
                return options;
            }
        }
    }
}